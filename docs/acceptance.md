# r3 acceptance gates

- Preserve existing sheet columns, records, UI and cipher semantics.
- Require authenticated server operations; no credentials in public repository.
- Confirm durable writes before clearing input; retain stable operation IDs across retries.
- Reject unknown actions and invalid values; validate lot identity and status.
- Preserve audit history, prevent stale master updates and stale lot responses.
- Test on a copy of production spreadsheet before deployment.
- Verify iPhone, PC, concurrent devices, lost responses and reload recovery.
- Do not merge until GAS deployment and end-to-end verification are complete.

Status: requirements only; implementation and tests pending. Cipher rounding and completion residual policy require business confirmation.
