# UTY Development Dashboard v4.3.0

Corrected K1 business logic:

- Only Active records are included.
- Dropped records are excluded.
- K1 Plan values blank, -, TBD, No need, N/A are excluded.
- Total K1 styles counts only active styles requiring K1.
- First-pass rate = Pass in round 1 / eligible K1 styles.
- Need re-fit = at least one Fail in any recorded round.
- Pass after re-fit = round 1 Fail, then later Pass.
- Still open = eligible K1 style with no Pass recorded.
- Average attempts uses submitted rounds only.
- Factory chart compares First-pass % and Need re-fit %.
- Global Season/Stage/Factory/Material/Model filters remain shared across tabs.
- Quality tab adds K1 Status and Attempts filters.
