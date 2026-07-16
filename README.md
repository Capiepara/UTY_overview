# UTY Development Dashboard v4.3.3

K1 logic correction:

- Only Active records are included.
- Dropped records are excluded.
- K1 Plan blank, -, TBD, No need, N/A are excluded.
- Explicit -, TBD, No need or N/A in K1 Round 1 Result means K1 is not required and is excluded.
- Only a truly blank Result cell is Waiting / Result Pending.
- Round 2 is evaluated only after Round 1 Fail.
- Round 3 is evaluated only after Round 1 and Round 2 both Fail.
- Pass stops the later-round evaluation.
- Global filters now appear above the Development / Quality tabs.

Raw file validation for SS27 + SMS + Active:
- 23 Active raw rows
- 1 row excluded because K1 Plan = "-"
- 11 K1-required rows with real Pass/Fail results
- 6 Pass round 1
- 5 Fail round 1 then Pass round 2
- 0 Waiting
- 0 Result pending
- 0 Still open
