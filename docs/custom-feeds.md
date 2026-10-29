# Custom legislative feeds

In **Setup & contacts → Step 1**, choose the legislature, check its bill identifiers, then enter a **Custom legislative RSS / Atom feed URL** and click **Test feed**. The test reports usable entries, skipped entries, and sample bill identifiers. Leave the URL blank to keep the built-in source. The setting is saved with the router setup and included in exported setup codes; changing the legislature clears the URL. Use only public URLs without credentials, tokens or private subscription keys.

Custom feeds replace the built-in source for new-filings checks and daily-report/status feed reads. They are fetched when used; entering a URL does not schedule a new server-side mirror. A failed custom source never falls back to a different state. Review checkpoints include the custom source and identifier configuration, so changing them requires a fresh review.

RSS 2.0, RSS/RDF and Atom XML are supported when entries identify bills using the configured prefixes (such as HB, SB, AB or S) in their titles or links and have valid publication/update dates. A time must include a timezone; a date-only value is accepted. Missing or ambiguous dates and unrecognized bill identifiers are counted as skipped. If any entries are skipped, complete-window review marking is disabled to avoid implying full coverage. Empty valid feeds are reported as empty. This is not a promise that every state's feed shape can be interpreted without an adapter.

Feeds must allow cross-origin browser access. Authentication-required or CORS-blocked sources need a public browser-readable feed or an administrator-managed mirror; no open proxy or credential forwarding is provided. Requests omit credentials and referrers and have a 20-second timeout. Documents larger than 4 MB and DTD/custom entities are rejected.

Custom entries are summary-only. State-specific full-text extraction, amendment-parent enrichment, statutory mappings, comprehensive status tracking and fiscal deadlines are not inferred from an RSS URL. Review full source documents and configure division routing and deadlines for the jurisdiction.
