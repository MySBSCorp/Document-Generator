# Deploying Document Generator to AWS

## Recommended architecture

This application is fully static after `npm run build`; OCR, authentication, and document generation run in the browser. It does not require EC2, a container, or a server-side application.

For production, use:

```text
Browser -> CloudFront (HTTPS) -> private S3 bucket -> dist/ files
```

S3 website hosting by itself works only for a non-production preview. S3 website endpoints do not support HTTPS, while Microsoft Entra SPA redirect URIs require HTTPS for non-localhost addresses. CloudFront supplies HTTPS, caching, and private access to S3. Do not put secrets in this application: every value included in a frontend build is visible to users.

The GitHub Actions workflow in `.github/workflows/deploy-s3-cloudfront.yml` builds on every push to `main`, uploads `dist/`, and invalidates CloudFront. It can also be run manually.

## Prerequisites

- An AWS account and permission to manage S3, CloudFront, and IAM
- Access to this repository's GitHub settings
- A Microsoft Entra app registration configured as a **Single-page application (SPA)**
- Node.js 22 for local verification (Vite 8 requires Node.js 20.19+ or 22.12+)
- AWS CLI v2 if you want to run the optional command-line checks

Choose values for these placeholders:

```text
AWS_ACCOUNT_ID=123456789012
AWS_REGION=us-east-1
S3_BUCKET=document-generator-production
GITHUB_ORG=MySBSCorp
GITHUB_REPO=Document-Generator
```

S3 bucket names are globally unique, so adjust the example name.

## 1. Verify the production build locally

Create a local `.env` file (it is gitignored):

```dotenv
VITE_MSAL_CLIENT_ID=00000000-0000-0000-0000-000000000000
VITE_MSAL_TENANT_ID=00000000-0000-0000-0000-000000000000
# Set this after the CloudFront hostname or custom domain is known.
VITE_MSAL_REDIRECT_URI=https://docs.example.com
```

Then build and preview:

```bash
npm ci
npm run build
npm run preview
```

Confirm that sign-in, W-9 upload/OCR, document generation, and downloads work. Vite places the deployable output in `dist/`.

## 2. Create a private S3 bucket

In **AWS Console -> S3 -> Create bucket**:

1. Select the deployment region.
2. Enter the globally unique bucket name.
3. Leave **Block all public access** enabled.
4. Enable bucket versioning if rollback/audit history is useful.
5. Enable default encryption (SSE-S3 is sufficient unless company policy requires KMS).
6. Create the bucket.

Do **not** enable S3 static website hosting for the recommended CloudFront setup. CloudFront will use the bucket's REST endpoint through Origin Access Control (OAC).

## 3. Create the CloudFront distribution

In **AWS Console -> CloudFront -> Create distribution**:

1. Choose the S3 bucket as the origin (not its website endpoint).
2. Create and attach a new **Origin Access Control**.
3. Choose **Redirect HTTP to HTTPS** for the viewer protocol policy.
4. Allow only `GET` and `HEAD` methods.
5. Use `index.html` as the default root object.
6. Use the managed **CachingOptimized** cache policy.
7. Create the distribution and copy its distribution ID and hostname, such as `d123example.cloudfront.net`.
8. When prompted, let CloudFront update the S3 bucket policy. If it does not, copy the OAC bucket-policy example shown by the console into **S3 -> Permissions -> Bucket policy**.

Wait until the distribution status is **Deployed**.

### Optional custom domain

For a domain such as `docs.example.com`:

1. Request or import an ACM certificate in **us-east-1**. CloudFront requires its viewer certificate in this region.
2. Add `docs.example.com` as an alternate domain name on the distribution and select the certificate.
3. Create a DNS alias/CNAME from `docs.example.com` to the CloudFront hostname.
4. Use `https://docs.example.com` as `VITE_MSAL_REDIRECT_URI`.

## 4. Configure Microsoft Entra ID

In **Microsoft Entra admin center -> App registrations -> your app -> Authentication**:

1. Add a **Single-page application** redirect URI matching the exact production origin:
   - `https://d123example.cloudfront.net`, or
   - `https://docs.example.com` when using a custom domain.
2. Match the scheme, hostname, port, and trailing slash exactly with `VITE_MSAL_REDIRECT_URI`.
3. Do not create or put a client secret in GitHub or frontend code. This SPA uses Authorization Code Flow with PKCE.

See `AUTH.md` for the full Entra configuration and access-control model.

## 5. Create a GitHub Actions AWS role (OIDC)

OIDC lets GitHub obtain short-lived AWS credentials; no permanent AWS access key is stored in GitHub.

First, create the GitHub OIDC identity provider in IAM if the AWS account does not already have one:

- Provider URL: `https://token.actions.githubusercontent.com`
- Audience: `sts.amazonaws.com`

Create an IAM role named `GitHubDocumentGeneratorDeploy` with this trust policy. Replace the account/repository values if needed:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": {
        "Federated": "arn:aws:iam::AWS_ACCOUNT_ID:oidc-provider/token.actions.githubusercontent.com"
      },
      "Action": "sts:AssumeRoleWithWebIdentity",
      "Condition": {
        "StringEquals": {
          "token.actions.githubusercontent.com:aud": "sts.amazonaws.com",
          "token.actions.githubusercontent.com:sub": "repo:MySBSCorp/Document-Generator:ref:refs/heads/main"
        }
      }
    }
  ]
}
```

Attach this least-privilege permissions policy, replacing the bucket name and distribution ID:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "ListDeploymentBucket",
      "Effect": "Allow",
      "Action": ["s3:ListBucket"],
      "Resource": "arn:aws:s3:::document-generator-production"
    },
    {
      "Sid": "PublishStaticFiles",
      "Effect": "Allow",
      "Action": ["s3:DeleteObject", "s3:GetObject", "s3:PutObject"],
      "Resource": "arn:aws:s3:::document-generator-production/*"
    },
    {
      "Sid": "InvalidateCloudFront",
      "Effect": "Allow",
      "Action": "cloudfront:CreateInvalidation",
      "Resource": "arn:aws:cloudfront::AWS_ACCOUNT_ID:distribution/CLOUDFRONT_DISTRIBUTION_ID"
    }
  ]
}
```

Copy the role ARN, for example `arn:aws:iam::123456789012:role/GitHubDocumentGeneratorDeploy`.

## 6. Configure the GitHub production environment

In **GitHub -> repository -> Settings -> Environments**, create an environment named `production`. Add deployment approvals or branch protection if required by your release process.

Add these **environment variables**:

| Name | Example |
| --- | --- |
| `AWS_REGION` | `us-east-1` |
| `S3_BUCKET` | `document-generator-production` |
| `CLOUDFRONT_DISTRIBUTION_ID` | `E123EXAMPLE` |
| `VITE_MSAL_CLIENT_ID` | Entra application (client) ID |
| `VITE_MSAL_TENANT_ID` | Entra directory (tenant) ID |
| `VITE_MSAL_REDIRECT_URI` | `https://docs.example.com` |

Add this **environment secret**:

| Name | Value |
| --- | --- |
| `AWS_ROLE_ARN` | ARN of `GitHubDocumentGeneratorDeploy` |

The Entra IDs are public identifiers, not secrets, but environment variables keep environment-specific configuration out of source control.

## 7. Deploy

Merge or push a change to `main`, or run **Actions -> Deploy to S3 and CloudFront -> Run workflow**.

The workflow will:

1. Install the locked dependencies with `npm ci`.
2. Build the app using the production Entra values.
3. Assume the AWS role using GitHub OIDC.
4. Synchronize `dist/` to S3 and delete files no longer in the build.
5. Mark `index.html` as non-cacheable and invalidate the CloudFront distribution.

Open the CloudFront/custom-domain URL after the job succeeds. The first deployment can take several minutes to propagate.

## Verification checklist

- The site loads over HTTPS and HTTP redirects to HTTPS.
- Browser refresh at `/` returns the application.
- Microsoft sign-in returns to the same production origin without `AADSTS50011`.
- The logo and MSA template load from `/assets/`.
- W-9 OCR and document generation work.
- Sign-out returns to the production site.
- The S3 bucket is not publicly accessible directly.
- A second deployment updates the site and removes deleted build files.

## Rollback

The simplest rollback is to revert the bad commit and push the revert to `main`; the workflow redeploys the previous source. If S3 versioning is enabled, objects can also be restored from prior versions, but a Git revert is preferred because it keeps source control and production synchronized.

## S3-only preview (not recommended for production)

If authentication is disabled and HTTP is acceptable, enable S3 static website hosting, set `index.html` as the index document, make the published objects publicly readable, and upload `dist/`:

```bash
aws s3 sync dist/ s3://YOUR_BUCKET --delete
```

Use the S3 **website endpoint**, not the bucket REST URL. This option has no HTTPS support and therefore is not suitable for this app's production Entra sign-in. Keep the bucket private and use CloudFront for production.

## Troubleshooting

- **`AADSTS50011`**: the Entra SPA redirect URI and `VITE_MSAL_REDIRECT_URI` do not match exactly.
- **CloudFront 403**: confirm the origin uses the S3 REST endpoint, OAC is attached, and the bucket policy permits that distribution.
- **Old UI after deploy**: wait for the invalidation to complete, then hard-refresh. Confirm the workflow targeted the expected distribution.
- **Missing assets**: inspect the workflow's build output and verify `dist/assets/logo.jpg` and `dist/assets/msa-template.docx` exist.
- **GitHub cannot assume the role**: confirm the OIDC provider, repository name, branch in the role trust policy, and `AWS_ROLE_ARN`.

