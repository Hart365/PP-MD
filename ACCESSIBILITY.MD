# Accessibility

## People first, not a tick in a box

Accessibility matters to me. PP-MD should help people document their Power Platform solutions, not put another barrier in their way.

That means thinking about how people actually use the app. A mouse isn't everyone's starting point. Neither is being able to see a diagram, distinguish a colour, or work through a busy screen without difficulty. Accessibility belongs in what I deliver, not on a list of things to consider afterwards.

I'm still learning, and there will be things I've missed. If something gets in your way, I want to hear about it. You don't need to be an Accessibility expert to tell me that something doesn't work for you.

## What I'm working towards

The target for PP-MD's interface is [WCAG 2.2 Level AA](https://www.w3.org/TR/WCAG22/). In practical terms, that means:

- Being able to complete tasks with a keyboard, with a visible and sensible focus order.
- Giving controls clear names, and making progress, errors, and changes understandable to screen-reader users.
- Making text and controls readable in light and dark themes, without relying on colour alone.
- Supporting zoom, smaller windows, reduced motion, and different ways of interacting.
- Making the documentation and exported content usable too. An accessible interface is only part of the job.

The app already includes keyboard file selection alongside drag and drop, focus styles, status announcements, light and dark themes, reduced-motion support, document search, and diagram source fallbacks. Export and Options dropdowns have full-row hover highlights, keyboard-focus outlines, and dismissal with Escape or an outside click. There are automated checks for selected text and interface colour combinations in both themes, including highlighted dropdown states.

Those are useful foundations. They are **not a declaration of full WCAG conformance**, and a passing automated check doesn't tell me whether you can actually finish your task.

## Known limitations and gaps

This statement describes the current implementation and its verification gaps. It isn't a completed Accessibility audit.

### Diagrams

Mermaid diagrams are visual, and a diagram name or its source code isn't the same as a useful description of the relationships it shows. Source text and the surrounding documentation may help, but they aren't guaranteed to provide an equivalent experience.

Try the raw Markdown view or export the Markdown to an editor you use comfortably. Please report any information you can only get from a diagram.

Diagrams remain inline unless you choose a separate diagrams document. When separated, companion documents appear in the sidebar's Diagrams section. Turning off Include diagrams removes them from the generated documentation.

### PDF exports

The searchable PDF export contains text, but it isn't verified as a tagged, accessible PDF or as PDF/UA conformant. Visual headings and tables shouldn't be assumed to provide the structure a screen reader needs. Diagrams are embedded as images, and the PDF fonts also replace or omit some unsupported characters.

The desktop app also retains a tagged-PDF export option using Chromium's PDF output. Requesting tags and a document outline doesn't by itself establish correct reading order, equivalent diagram descriptions, or PDF/UA conformance. That output still needs to be checked with suitable assistive technology.

Markdown or HTML export may be a more useful starting point, depending on your tools. These formats aren't a guarantee of Accessibility either; check the result before sharing it with others.

### Large documents and other exports

Wide metadata tables and complex diagrams can be difficult to navigate or magnify. Excel exports and generated HTML haven't been comprehensively verified with assistive technology. Imported Markdown and descriptions from your solution can also bring their own Accessibility problems.

### Desktop application and assistive technology

PP-MD is a standalone Windows desktop application built with Electron. It processes exported solution ZIP files locally, without uploading them. Browser-based local development is also available. Native Windows file dialogs, operating-system settings, and the applications used to open exports are part of your experience too.

The repository includes automated Chromium accessibility checks and regression checks for keyboard interaction, selected light/dark theme states, narrow-screen reflow, and Windows forced-colours rendering. Those checks cover particular scenarios, not every screen, export, or assistive-technology combination.

End-to-end screen-reader behaviour in the packaged Windows app, keyboard focus around native file dialogs, 200% text resizing, and 400% zoom still need wider manual validation. There isn't a verified Windows, Electron, browser, and assistive-technology support matrix to publish yet.

If you're unsure whether a barrier belongs to PP-MD or another part of your setup, report it here anyway. I can help work out where it needs to go.

## Found a barrier? Please raise an Issue

1. Open [a new Accessibility barrier report](https://github.com/Hart365/PP-MD/issues/new?template=accessibility.yml). You'll need to sign in to GitHub.
2. Give it a short title describing the problem, such as "Keyboard focus disappears after closing export options".
3. Tell me what you were trying to do, what happened, and what you needed to happen instead.
4. If you can, add the PP-MD version, whether you're using the portable or installed app (or a browser for local development), Windows or browser version, and any relevant assistive technology or settings.
5. Submit the Issue. You can return to it later to add information or answer follow-up questions.

Steps to reproduce, the impact on your task, and any workaround are helpful. Screenshots, recordings, and technical details are welcome **if they help and you're comfortable sharing them**, but they aren't required. A plain-language description is enough to start. You don't need a WCAG reference, a diagnosis, or an explanation of your disability.

GitHub Issues are public. Please don't attach customer solution ZIPs, generated documents containing sensitive metadata, credentials, personal information, or unredacted screenshots. Use a small, made-up example where possible. For a suspected security vulnerability, follow the [Security policy](SECURITY.md) instead of opening a public Issue.

### If GitHub itself is a barrier

If you can't use the Issue form, try [a blank Issue](https://github.com/Hart365/PP-MD/issues/new). If GitHub isn't usable for you, contact me, Mike Hartley, through [LinkedIn](https://www.linkedin.com/in/mikebhartley/). Let me know your preferred way to continue the conversation. I'll ask before putting anything you share privately into a public Issue.

## What happens next?

I'll review reports, ask for clarification where needed, and use the Issue to record progress, a workaround, or a proposed fix. Problems that stop someone completing a task need particular attention, especially where there isn't a workable alternative.

This is a community project, so I can't promise a response or fix within a set time. I can commit to taking barriers seriously, being honest about limitations, and learning from the people experiencing them. A report doesn't need to arrive with a solution attached.

Contributors can find practical checks in the [contributing guide](CONTRIBUTING.md#accessibility). Thanks for helping make this better.