# Bill Summarizer — Copilot agent setup (requires M365 Copilot licenses)

Renee's ask: "Copilot summarizes how the bill is written, posted in the
channel for all to see." With M365 Copilot licenses (Iowa DOT has them),
you can build this **without any additional purchase** using declarative
agents in Copilot's built-in Agent Builder. Internal documents only —
grounding is limited to sources you pick.

## Build it (10 minutes, no code)

1. Open **Microsoft 365 Copilot** (app or m365.cloud.microsoft) → **Create
   agent** (Agent Builder) → **Configure** tab.
2. Name: `Bill Summarizer`.
3. Instructions — paste:

   ```
   You summarize Iowa legislative bills and amendments for Iowa DOT staff.
   When given a bill number or bill text, produce:
   1) A 3–5 sentence plain-language summary of what the bill does.
   2) A "DOT relevance" line: which DOT functions it touches
      (motor vehicle, highways, transportation development, funding,
      personnel, procurement) and why.
   3) Key sections/Code chapters amended, as a short list.
   4) Effective dates and any successor/predecessor bills mentioned.
   Always state that this is an AI summary and staff must read the bill
   text before commenting. Do not offer legal conclusions.
   ```

4. Knowledge sources: add the **legislative SharePoint site** (the one with
   the routing/tracker lists and any saved bill PDFs/DOCX). Optionally add
   the Teams channel. Do NOT enable web grounding if the requirement is
   internal-documents-only.
5. Create → **share the agent with the legislative Teams channel's team**.

## Use it in the channel

Staff type `@Bill Summarizer summarize HF935` in the channel (after adding
the agent to the team). The summary posts in-thread — visible to all, which
is exactly the requested behavior. For bills not yet saved internally,
paste the bill text into the prompt, or save the PDF to the SharePoint
site first (the file-renaming macro's output folder is a natural fit).

## Notes

- Declarative agents are included with M365 Copilot licensing; Copilot
  Studio (paid) is only needed for autonomous/scheduled agents.
- The Router's Teams posts already include the LSA's official bill
  description, so every post carries a neutral baseline summary even
  without invoking the agent.
- Governance: DOT's Copilot admin may need to approve agent sharing;
  loop in the Co-Pilot Teams chat Renee mentioned.
