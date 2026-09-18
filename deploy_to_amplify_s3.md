# Deploy Document Generator with AWS Amplify Hosting and S3

## Recommendation

Deploy this application with **AWS Amplify Hosting connected directly to GitHub**.

This application is a static Vite frontend. Microsoft Entra authentication, W-9 OCR, and document generation all execute in the user's browser, so it does not need EC2, nginx, a container, or a continuously running server.

The production URL will be:

```text
https://w9.mysbscorp.com
```

Amplify provides the HTTPS certificate, hosting, managed CDN, custom-domain connection, and deployment on every push. AWS implements Amplify Hosting with managed storage and CloudFront, but there is no CloudFront distribution for this team to create, configure, invalidate, patch, or monitor directly.

An independent S3 website without Amplify or CloudFront is **not suitable for this application**. S3 website endpoints support HTTP only, and Microsoft Entra requires HTTPS for non-localhost SPA redirect URIs. Slow rendering would be acceptable, but it does not solve the HTTPS and certificate requirement.

AWS references:

- [AWS Amplify pricing](https://aws.amazon.com/amplify/pricing/)
- [AWS: Deploying a static website to Amplify from S3](https://docs.aws.amazon.com/amplify/latest/userguide/deploy-website-from-s3.html)
- [AWS: S3 website endpoints do not support HTTPS](https://docs.aws.amazon.com/AmazonS3/latest/userguide/WebsiteEndpoints.html)
- [AWS: Updating GoDaddy DNS records for Amplify](https://docs.aws.amazon.com/amplify/latest/userguide/to-add-a-custom-domain-managed-by-godaddy.html)

## What “Amplify and S3” means

There are two supported deployment models:

1. **GitHub -> Amplify Hosting (recommended):** Amplify checks out the repository, runs `npm ci` and `npm run build`, and hosts `dist/`. No separately managed S3 bucket is needed.
2. **GitHub Actions -> your S3 bucket -> Amplify Hosting:** the workflow builds and uploads `dist/` to a general-purpose S3 bucket, and Amplify publishes that bucket to its managed HTTPS hosting. This adds a bucket and deployment plumbing without providing a meaningful benefit for this project.

Use the first model unless company policy specifically requires build artifacts to pass through a company-owned S3 bucket.

## Cost analysis

### Assumptions

The estimates below are in USD, exclude taxes and AWS Support plans, and intentionally ignore promotional/free-tier credits so the comparison remains useful after those credits expire.

- Current production build: approximately **4.4 MB**
- 100 complete application loads per month
- Approximately 0.44 GB served per month (`4.4 MB x 100`)
- Four deployments per month
- Two Amplify build minutes per deployment
- One EC2 instance running continuously for 730 hours per month
- EC2 comparison uses Linux `t3.micro` pricing shown by AWS for US East (N. Virginia)
- EC2 disk comparison uses an 8 GB gp3 volume in a region priced at $0.08/GB-month
- One public IPv4 address for EC2
- Low traffic remains within AWS's first 100 GB/month shared internet data-transfer allowance where applicable

Actual prices vary by region, account, traffic, build frequency, and AWS pricing changes. Confirm the selected region in the [AWS Pricing Calculator](https://calculator.aws/) before approval.

### Estimated monthly cost

| Component | Amplify from GitHub | S3 source + Amplify | EC2 `t3.micro` |
| --- | ---: | ---: | ---: |
| Build/deployment compute | `8 min x $0.01` = **$0.08** | GitHub Actions cost depends on the GitHub plan; no Amplify build required for prebuilt files | Included in instance administration |
| Hosting storage | Well under **$0.01** | Amplify storage under **$0.01**, plus S3 staging under **$0.01** | 8 GB gp3: about **$0.64** |
| Data served | `0.44 GB x $0.15` = about **$0.07** | About **$0.07** through Amplify | Usually **$0** at this usage if covered by AWS's shared 100 GB allowance |
| Server compute | **$0** | **$0** | `730 x $0.0104` = about **$7.59** |
| Public IPv4 | **$0** dedicated-IP charge | **$0** dedicated-IP charge | `730 x $0.005` = about **$3.65** |
| HTTPS certificate | Included | Included | Let's Encrypt can be free, but must be installed and renewed |
| Approximate total | **$0.15/month** | **$0.08-$0.20/month**, plus any GitHub Actions charge | **$11.88/month** |

These estimates use AWS's published Amplify rates of $0.01 per standard build minute, $0.023 per GB-month stored, and $0.15 per GB served. AWS lists `t3.micro` at $0.0104/hour, example gp3 storage at $0.08/GB-month, and public IPv4 addresses at $0.005/hour.

At this app's usage level, the exact Amplify total is less important than the order of magnitude: it should remain well below $1/month unless deployments or traffic increase substantially. EC2 incurs its fixed monthly cost even if nobody opens the application.

### Operational comparison

| Concern | Amplify Hosting | EC2 |
| --- | --- | --- |
| Server patching | None | Required |
| nginx configuration | None | Required |
| HTTPS certificate | Managed and automatically renewed | Configure and monitor renewal |
| Deployment | Automatic from Git | Custom workflow/SSH/rsync |
| Scaling and availability | Managed | Single instance is a single point of failure |
| Idle cost | Near zero | Instance, disk, and IP billed continuously |
| Rollback | Redeploy a prior commit/build | Maintain and switch releases manually |

**Decision:** use Git-connected Amplify Hosting. It is cheaper and has materially lower operational risk for this static, rarely accessed application.

## Prerequisites

- AWS account access to Amplify Hosting
- Administrator or appropriate access to `MySBSCorp/Document-Generator` on GitHub
- GoDaddy DNS access for `mysbscorp.com`
- Microsoft Entra app-registration access
- The `w9.mysbscorp.com` DNS name must be available to point at Amplify
- Production changes should be merged to `main`

## Option A — Deploy from GitHub to Amplify (recommended)

### 1. Verify the build

Run locally:

```bash
npm ci
npm run build
```

The build must complete successfully and create `dist/index.html`. The current build also includes `dist/assets/logo.jpg`, `dist/assets/msa-template.docx`, and the PDF.js worker.

Do not upload the repository root as the website. Only the generated `dist/` directory is deployable.

### 2. Create the Amplify application

1. Sign in to the AWS Console.
2. Open **AWS Amplify**.
3. Choose **Create new app** or **Deploy an app**.
4. Select **GitHub** as the source-code provider.
5. Authorize the AWS Amplify GitHub App when prompted.
6. Select the `MySBSCorp/Document-Generator` repository.
7. Select the `main` branch for production.
8. Choose a descriptive app name such as `document-generator-production`.
9. Keep automatic deployments enabled. Every push or merge to `main` will trigger a new build and deployment.

### 3. Configure the Amplify build

On the build-settings page, use:

```yaml
version: 1
frontend:
  phases:
    preBuild:
      commands:
        - nvm use 22
        - npm ci
    build:
      commands:
        - npm run build
  artifacts:
    baseDirectory: dist
    files:
      - '**/*'
  cache:
    paths:
      - ~/.npm/**/*
```

Important values:

```text
Build command: npm run build
Output directory: dist
Node.js version: 22
```

The build currently warns about large JavaScript chunks. That warning does not fail deployment and is expected because OCR, PDF, and Word-generation libraries are shipped to the browser.

### 4. Add environment variables

Before starting the production build, open the app's **Hosting -> Environment variables** or **App settings -> Environment variables** page and add:

| Variable | Value |
| --- | --- |
| `VITE_MSAL_CLIENT_ID` | Microsoft Entra application/client ID |
| `VITE_MSAL_TENANT_ID` | Microsoft Entra directory/tenant ID |
| `VITE_MSAL_REDIRECT_URI` | `https://w9.mysbscorp.com` |

The client and tenant IDs are public identifiers, not confidential credentials. Never add a Microsoft client secret to this frontend or to a `VITE_` variable; Vite embeds those variables into browser-readable JavaScript.

If Amplify requires the first deployment before domain setup, initially set `VITE_MSAL_REDIRECT_URI` to the generated Amplify URL, deploy, connect the custom domain, change it to `https://w9.mysbscorp.com`, and redeploy.

### 5. Deploy the first build

1. Review the branch, build settings, and environment variables.
2. Choose **Save and deploy**.
3. Wait for **Provision**, **Build**, **Deploy**, and **Verify** to succeed.
4. Open the temporary URL, similar to `https://main.d123example.amplifyapp.com`.
5. Confirm that the page and static assets load. Sign-in will work only after that exact temporary URL or the final custom URL is registered in Entra.

If the build fails, open the Amplify build log. The most common causes are an incorrect output directory, missing environment variables, or a Node version older than the version required by Vite.

## Connect `w9.mysbscorp.com` from GoDaddy

### 6. Start the custom-domain connection in Amplify

1. In the Amplify app, open **Hosting -> Custom domains** or **Domain management**.
2. Choose **Add domain**.
3. Enter `mysbscorp.com`.
4. Choose **Manual configuration** because DNS remains at GoDaddy.
5. Map the `w9` subdomain to the production `main` branch.
6. Use the Amplify-managed certificate. There is no additional certificate charge.
7. Continue until Amplify displays its DNS records.

Amplify normally shows two CNAME records:

- A certificate/domain-verification CNAME whose name begins with an underscore.
- A `w9` CNAME pointing to an AWS hostname, commonly ending in `cloudfront.net`.

Always copy the exact values shown for this Amplify app; do not copy example hashes or hostnames from this document.

### 7. Update DNS in GoDaddy

1. Sign in to GoDaddy.
2. Open **My Products -> Domains -> mysbscorp.com -> DNS -> Manage DNS**.
3. Find any existing record with host/name `w9`.
4. Record its current value for rollback.
5. If `w9` points to the old EC2 deployment, edit or replace that record only when ready to cut over.
6. Add the certificate-verification record supplied by Amplify:

```text
Type:  CNAME
Name:  the underscore-prefixed host supplied by Amplify
Value: the acm-validations.aws value supplied by Amplify
TTL:   600 seconds or GoDaddy default
```

GoDaddy usually appends `.mysbscorp.com` automatically. For the **Name**, enter only the host portion shown before `.mysbscorp.com`.

7. Add or update the application record:

```text
Type:  CNAME
Name:  w9
Value: the target hostname supplied by Amplify
TTL:   600 seconds or GoDaddy default
```

Do not include `https://`, `/`, or any path in a DNS value.

8. Save the records.
9. Do not delete the underscore-prefixed certificate-validation CNAME after activation. Amplify needs it to renew the certificate automatically.
10. Return to Amplify and wait until the domain status is **Available**. DNS and certificate validation may take from several minutes to, in unusual cases, 48 hours.

If GoDaddy already has a CNAME named `w9`, it must be edited rather than duplicated. A DNS name cannot simultaneously have a CNAME and conflicting A/AAAA records.

### 8. Configure Microsoft Entra

In **Microsoft Entra admin center -> App registrations -> the Document Generator app -> Authentication**:

1. Add a platform if necessary and choose **Single-page application**.
2. Add this exact redirect URI:

```text
https://w9.mysbscorp.com
```

3. Remove the old production redirect URI only after the new deployment is verified and rollback is no longer needed.
4. Leave implicit-grant token options disabled; this application uses authorization code flow with PKCE.
5. Ensure the intended employees or groups are assigned to the enterprise application if assignment is required.

The URI must match the protocol and hostname exactly. Do not register `http://w9.mysbscorp.com`, `https://www.mysbscorp.com/w9`, or a trailing path unless the application is deliberately configured to use it.

### 9. Redeploy after domain configuration

If `VITE_MSAL_REDIRECT_URI` was initially set to the Amplify URL:

1. Change it to `https://w9.mysbscorp.com`.
2. Choose **Redeploy this version**, or push a new commit to `main`.
3. Wait for the deployment to finish.

### 10. Verify production

Open `https://w9.mysbscorp.com` in a private/incognito browser window and verify:

- HTTPS is valid with no certificate warning.
- Microsoft sign-in completes and returns to `https://w9.mysbscorp.com`.
- An unassigned Entra user is denied according to company policy.
- The logo loads.
- W-9 PDF/image upload and browser-side OCR work.
- The MSA template loads and Word/PDF generation succeeds.
- Generated files download correctly.
- Sign-out returns to the production URL.
- A push to `main` automatically triggers a new Amplify deployment.

This is an internal app by identity policy, not by network location. The URL is publicly reachable, but Entra controls who can use the application. Frontend source files and public static assets must not contain secrets or confidential server credentials.

## Option B — Use your own S3 bucket as Amplify's deployment source

Use this only if an organizational requirement calls for keeping a copy of each compiled deployment in an S3 bucket. Users still access the app through Amplify HTTPS hosting, not through the S3 website endpoint.

### 1. Create the artifact bucket

1. Open **Amazon S3 -> Create bucket**.
2. Use a globally unique name, such as `mysbscorp-document-generator-artifacts`.
3. Choose an AWS Region in which Amplify Hosting is available.
4. Keep **Block all public access** enabled.
5. Enable bucket versioning for rollback/audit history.
6. Keep default encryption enabled.
7. Do not enable S3 static website hosting.

### 2. Build and upload artifacts

Build locally or in GitHub Actions:

```bash
npm ci
npm run build
aws s3 sync dist/ s3://mysbscorp-document-generator-artifacts/current/ --delete
```

Only `dist/` should be uploaded. Keep source files, `.env` files, tests, and Git metadata out of the bucket.

### 3. Create an Amplify deployment from S3

1. Open the Amplify console.
2. Choose **Create new app -> Deploy without Git** or the available **Amazon S3** deployment option.
3. Select **Browse S3**.
4. Select the bucket and the `current/` prefix containing `index.html`.
5. Choose an app name such as `document-generator-production`.
6. Save and deploy.
7. Complete the GoDaddy and Entra steps above.

AWS supports deployments from a general-purpose S3 bucket in the same AWS account. The bucket must be in a Region where Amplify is available. Charges for the hosted website follow the Amplify pricing model; the source bucket adds only normal S3 storage and request charges.

### 4. Automating S3-source deployments

S3-source deployment does not provide the same simple Git build pipeline as connecting Amplify directly to GitHub. A complete automated flow must:

1. Build on a push to `main`.
2. Upload `dist/` to the S3 source prefix.
3. Trigger/update the Amplify deployment.
4. Verify the deployment.

Because Git-connected Amplify already performs these steps with less IAM and pipeline configuration, use Option A unless the S3 source bucket is mandatory.

## Why direct S3-only hosting is not used

It is technically possible to publish the build at an address such as:

```text
http://w9.mysbscorp.com
```

by naming a public S3 website bucket `w9.mysbscorp.com` and pointing the GoDaddy CNAME to the S3 website endpoint. It is not acceptable for this application because:

1. S3 website endpoints do not support HTTPS.
2. Microsoft Entra does not allow HTTP SPA redirect URIs for non-localhost addresses.
3. The bucket and website objects must be public.
4. Browsers would label the application insecure.
5. W-9 data is handled in the browser, so transport security must not be weakened even though processing is client-side.

The S3 REST endpoint supports HTTPS for its AWS hostname, but it is not a replacement for static website hosting on `w9.mysbscorp.com`: it does not provide a custom-domain certificate or normal website index behavior.

Therefore, the supported low-operations choices are:

- **Amplify Hosting**, with its AWS-managed HTTPS/CDN layer; or
- A manually managed **S3 + CloudFront** deployment.

For this app, choose Amplify.

## Rollback

### Application rollback

1. Open the Amplify app's deployment history.
2. Locate the last known-good deployment.
3. Choose **Redeploy this version**, if available, or revert the bad Git commit and push the revert to `main`.
4. Run the production verification checklist.

Reverting in Git is preferred because the source repository and production remain synchronized.

### DNS rollback

If the Amplify domain connection fails during migration:

1. Restore the previous `w9` DNS record value recorded before cutover.
2. Wait for its TTL to expire.
3. Keep the old Entra redirect URI until migration is complete.

Do not delete the old EC2 deployment until Amplify, DNS, authentication, uploads, and generated downloads have all been verified.

## Troubleshooting

- **`AADSTS50011`**: the Entra SPA redirect URI and `VITE_MSAL_REDIRECT_URI` do not match exactly.
- **Amplify build cannot find output**: set the artifact `baseDirectory` to `dist`.
- **Vite fails due to Node version**: confirm `nvm use 22` executes before `npm ci` and `npm run build`.
- **Custom domain remains pending**: compare both Amplify CNAME records with GoDaddy and ensure GoDaddy did not append the domain twice.
- **CNAME already exists**: edit the existing `w9` record instead of adding a duplicate, and confirm the domain is not attached to another Amplify app or CloudFront distribution.
- **Certificate does not issue**: check the underscore-prefixed ACM validation CNAME and any restrictive CAA records.
- **Old content appears**: verify the latest Amplify job succeeded, then hard-refresh or test in an incognito window.
- **Missing logo/template/PDF worker**: inspect the successful build artifacts and confirm the expected files exist under `dist/assets/`.
- **Temporary Amplify URL works but custom domain returns 404**: confirm `w9.mysbscorp.com` is mapped to the `main` branch in Amplify and the GoDaddy CNAME target matches Amplify exactly.

