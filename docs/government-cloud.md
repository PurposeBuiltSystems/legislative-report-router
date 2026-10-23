# Government cloud configuration

Graph and sign-in endpoints are selected by configuration. This does not
certify operation in a government cloud: validate authentication, required
Graph APIs, hosting policy and a complete publication workflow in the target
tenant before deployment.

The Settings → "Cloud environment" selector maps:

| Environment | Graph host | Sign-in authority |
| --- | --- | --- |
| Commercial | `https://graph.microsoft.com` | `https://login.microsoftonline.com` |
| GCC | `https://graph.microsoft.com` | `https://login.microsoftonline.com` |
| GCC High | `https://graph.microsoft.us` | `https://login.microsoftonline.us` |
| DoD | `https://dod-graph.microsoft.us` | `https://login.microsoftonline.us` |

Notes:

- GCC runs on the commercial endpoints (it is a compliance boundary, not an
  endpoint boundary); GCC High and DoD use the `.us` sovereign endpoints.
- GCC High/DoD require a separate application registration in the sovereign
  cloud. A commercial multitenant registration cannot span national clouds.
  For a separate sovereign deployment, set `CLIENT_ID` in `src/graph.js` to
  that registration's client ID, configure the appropriate SPA/broker redirect
  for its host, and update manifest URLs if hosting changes. Selecting a cloud
  in the public deployment does not replace its commercial registration.
  See [Microsoft's national-cloud authentication guidance](https://learn.microsoft.com/en-us/entra/identity-platform/authentication-national-cloud).
- Teams tag mention support and Graph feature parity can lag in sovereign
  clouds; if `mentions` with `mentioned.tag` is rejected, the post still
  succeeds without the mention and the audit row records the response.
