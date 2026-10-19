# ohio_public_schools
Ohio schools that are not a part of the other vendors

## SAM deployment

The deployable SAM template is at [template.yaml](template.yaml). It defines the four import Lambdas under [sam-src](sam-src) and uses the shared AWS SDK for pandas layer.

Typical commands:

```bash
sam build --use-container
sam deploy --guided
```

To reapply the bucket notifications for the existing raw bucket, use the helper in [scripts/README.md](scripts/README.md).
