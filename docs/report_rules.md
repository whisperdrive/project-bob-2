# Report reading rules

Read by the report agents on every run (bench/lessons.py): the table reader, the fact extractor, its
reviewer, and both sides of the review loop. Each rule has an ID the agents cite. Rules are about method only:
no names, figures or quotes from any engagement. Learned lessons (L1, L2, ...) live in out/lessons.json on each
machine; promote one here from the Report reference step when it has proved itself.

- **R1** [tables] Copy column headings with their currency and unit markers exactly as printed, including any currency prefix before the unit. _Why: a dropped prefix changes what the numbers mean and fails the text-layer check._
- **R2** [tables] Keep every value under the column it is printed in; leave blank cells empty instead of shifting later values left. _Why: a shifted row puts figures under the wrong year or scenario._
- **R3** [facts] Quote one continuous piece of the document; for a figure in a table, quote its row and name the column it sits under in basis. _Why: the check looks for the quote on the cited page as one piece._
- **R4** [facts] For a valuation range, put the preferred value in value_text and the ends in low_text and high_text; never compute a midpoint the report doesn't state. _Why: the reference must match what the report prints._
- **R5** [facts] When a figure appears more than once, use its main statement (executive summary or assumptions table) and record whether it is pre- or post-tax, nominal or real. _Why: the same rate is often restated on another basis elsewhere._
