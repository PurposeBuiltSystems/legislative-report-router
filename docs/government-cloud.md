# Government cloud configuration

No endpoint is hardcoded. The Settings → "Cloud environment" selector maps:

| Environment | Graph host | Sign-in authority |
| --- | --- | --- |
| Commercial | `https://graph.microsoft.com` | `https://login.microsoftonline.com` |
| GCC | `https://graph.microsoft.com` | `https://login.microsoftonline.com` |
| GCC High | `https://graph.microsoft.us` | `https://login.microsoftonline.us` |
| DoD | `https://dod-graph.microsoft.us` | `https://login.microsoftonline.us` |

Notes:

- GCC runs on the commercial endpoints (it is a compliance boundary, not an
  endpoint boundary); GCC High and DoD use the `.us` sovereign endpoints.
- For GCC High/DoD the Entra app registration must exist in (or be consented
  into) the sovereign tenant; SPA redirect URIs are cloud-agnostic (GitHub
  Pages), but agencies that disallow external hosting can serve the static
  files from any internal HTTPS host — update the manifest URLs, nothing
  else changes (there is no backend).
- Teams tag mention support and Graph feature parity can lag in sovereign
  clouds; if `mentions` with `mentioned.tag` is rejected, the post still
  succeeds without the mention and the audit row records the response.
