#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Deploy dist/ to the EC2 instance over SSH, atomically.
#
# Strategy: rsync into a fresh timestamped directory under releases/, then
# repoint the `current` symlink in one operation. nginx's root is the symlink,
# so the site flips from old to new between two requests — visitors never see
# a half-copied tree. Rollback is repointing the symlink at the previous
# release, which is why old releases are kept.
#
#   /var/www/docgen/
#     releases/20260819-143001/   <- this deploy
#     releases/20260819-101233/   <- previous, kept for rollback
#     current -> releases/20260819-143001
#
# No sudo is used anywhere. /var/www/docgen is owned by the deploy user (see
# DEPLOY.md step 4), and nginx needs no reload to follow a changed symlink.
#
# Config comes from scripts/deploy.env (gitignored, copy from
# deploy.env.example) or from the environment. Run via `npm run deploy`,
# which builds dist/ first.
# ---------------------------------------------------------------------------

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# shellcheck source=/dev/null
[[ -f "$SCRIPT_DIR/deploy.env" ]] && source "$SCRIPT_DIR/deploy.env"

DEPLOY_HOST="${DEPLOY_HOST:-}"
DEPLOY_USER="${DEPLOY_USER:-ec2-user}"
DEPLOY_KEY="${DEPLOY_KEY:-}"
DEPLOY_PORT="${DEPLOY_PORT:-22}"
REMOTE_ROOT="${REMOTE_ROOT:-/var/www/docgen}"
SITE_URL="${SITE_URL:-}"
KEEP_RELEASES="${KEEP_RELEASES:-5}"

RED=$'\033[31m'; GREEN=$'\033[32m'; YELLOW=$'\033[33m'; DIM=$'\033[2m'; OFF=$'\033[0m'
die() { echo "${RED}error:${OFF} $*" >&2; exit 1; }
step() { echo "${GREEN}==>${OFF} $*"; }

# --- preflight -------------------------------------------------------------

[[ -n "$DEPLOY_HOST" ]] || die "DEPLOY_HOST is not set. Copy scripts/deploy.env.example to scripts/deploy.env and fill it in."
[[ -d "$REPO_ROOT/dist" ]] || die "dist/ not found. Run 'npm run build' first (or use 'npm run deploy', which builds for you)."
[[ -f "$REPO_ROOT/dist/index.html" ]] || die "dist/index.html is missing — the build looks incomplete."
command -v rsync >/dev/null || die "rsync is not installed. macOS ships it; otherwise: brew install rsync"

SSH_OPTS=(-p "$DEPLOY_PORT" -o StrictHostKeyChecking=accept-new -o ConnectTimeout=10)
if [[ -n "$DEPLOY_KEY" ]]; then
  [[ -f "$DEPLOY_KEY" ]] || die "DEPLOY_KEY points at $DEPLOY_KEY, which does not exist."
  SSH_OPTS+=(-i "$DEPLOY_KEY")
fi
SSH=(ssh "${SSH_OPTS[@]}" "${DEPLOY_USER}@${DEPLOY_HOST}")

RELEASE="$(date +%Y%m%d-%H%M%S)"
REMOTE_RELEASE="${REMOTE_ROOT}/releases/${RELEASE}"

echo "${DIM}  host      ${DEPLOY_USER}@${DEPLOY_HOST}:${DEPLOY_PORT}"
echo "  remote    ${REMOTE_RELEASE}"
echo "  keeping   last ${KEEP_RELEASES} releases${OFF}"
echo ""

step "Checking SSH access and remote layout"
"${SSH[@]}" "test -d '${REMOTE_ROOT}/releases' && test -w '${REMOTE_ROOT}/releases'" \
  || die "${REMOTE_ROOT}/releases is missing or not writable by ${DEPLOY_USER}. Run DEPLOY.md step 4 on the instance first."

PREVIOUS="$("${SSH[@]}" "readlink -f '${REMOTE_ROOT}/current' 2>/dev/null || true")"
[[ -n "$PREVIOUS" ]] && echo "${DIM}  current release: ${PREVIOUS}${OFF}"

step "Uploading dist/ ($(du -sh "$REPO_ROOT/dist" | cut -f1))"
# --delete is safe here: the destination is a brand-new empty directory, so it
# only ever removes leftovers from an interrupted deploy of this same release.
rsync -az --delete --human-readable \
  --rsync-path="mkdir -p '${REMOTE_RELEASE}' && rsync" \
  -e "ssh ${SSH_OPTS[*]}" \
  "$REPO_ROOT/dist/" "${DEPLOY_USER}@${DEPLOY_HOST}:${REMOTE_RELEASE}/"

step "Verifying the upload landed intact"
"${SSH[@]}" "test -f '${REMOTE_RELEASE}/index.html' \
  && test -f '${REMOTE_RELEASE}/script.js' \
  && test -f '${REMOTE_RELEASE}/vendor/msal-browser-4.30.0.esm.js' \
  && test -f '${REMOTE_RELEASE}/vendor/tesseract/eng.traineddata.gz'" \
  || die "the uploaded release is missing expected files — the 'current' symlink was NOT changed, so the live site is untouched."

step "Switching 'current' to this release"
# -T stops ln from creating the link *inside* the existing directory; -f -n
# together replace the symlink in place rather than following it.
"${SSH[@]}" "ln -sfnT '${REMOTE_RELEASE}' '${REMOTE_ROOT}/current'"

step "Pruning old releases (keeping ${KEEP_RELEASES})"
# Never prune whatever `current` points at, no matter how old it is.
"${SSH[@]}" "cd '${REMOTE_ROOT}/releases' \
  && keep=\$(readlink -f '${REMOTE_ROOT}/current' | xargs -r basename) \
  && ls -1t | grep -v \"^\${keep}\$\" | tail -n +${KEEP_RELEASES} | xargs -r rm -rf"

if [[ -n "$SITE_URL" ]]; then
  step "Health check: ${SITE_URL}"
  # Confirm we get a 200 *and* that the body is really this app, not a
  # default nginx page or the other vhost on the box.
  code="$(curl -sS -o /tmp/docgen-healthcheck.html -w '%{http_code}' --max-time 20 "$SITE_URL" || echo 000)"
  if [[ "$code" != "200" ]]; then
    echo "${YELLOW}warning:${OFF} ${SITE_URL} returned HTTP ${code}."
    echo "  The files are deployed and 'current' was switched. This looks like a"
    echo "  DNS, TLS or nginx issue rather than a bad upload."
    echo "  On the instance:  sudo nginx -t && sudo tail -n 50 /var/log/nginx/docgen.error.log"
  elif ! grep -q 'Document Generator' /tmp/docgen-healthcheck.html; then
    echo "${YELLOW}warning:${OFF} ${SITE_URL} returned 200 but the body is not the Document Generator."
    echo "  Another vhost is probably answering for this hostname — check server_name"
    echo "  and which block is default_server."
  else
    echo "${GREEN}  200 OK and the page is the Document Generator.${OFF}"
  fi
  rm -f /tmp/docgen-healthcheck.html
else
  echo "${DIM}  (set SITE_URL in scripts/deploy.env to run an automatic health check)${OFF}"
fi

echo ""
echo "${GREEN}Deployed release ${RELEASE}.${OFF}"
if [[ -n "$PREVIOUS" && "$PREVIOUS" != "$REMOTE_RELEASE" ]]; then
  echo "Roll back with:"
  echo "  ssh ${DEPLOY_USER}@${DEPLOY_HOST} \"ln -sfnT '${PREVIOUS}' '${REMOTE_ROOT}/current'\""
fi
