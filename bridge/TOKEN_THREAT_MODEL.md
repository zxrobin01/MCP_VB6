# Token authentication - threat model and mechanism decision 

Status: **threat model FIRST, by author ruling** (ADR sec 0.5 clause 3 and the Lead's card).
This document decides WHO we defend against and WHAT they can reach; the mechanism at the end is
derived from that, not the other way round.  Written before any token code exists.

## 0. Where this sits

ADR `docs/design/TRANSPORT_NP_DUAL.md` sec 0.5 records the author's ruling (2026-10-15): token
authentication comes later, and until it existed **no non-loopback exposure was allowed**, the
remote per-machine shape could not be called deployable, and the mechanism had to be preceded by a
threat model.  **Since 2026-10-17 that rule is REFINED, not dropped**: remote access
now exists, but only as an EXPLICIT, fail-closed mode with its own hard prerequisites - section 10
is the binding text, and the default tier is unchanged.  The surfaces today:

- **Python bridge** (this file's component): ONE TCP port, 120-tool face **including write tools**
  (edit files, mutate the project, run code in the IDE), **no authentication**, loopback by default.
- **VB6 add-in pipe**: nonce-paired name + SDDL; the ADR notes cross-user negative tests were never done.
- **VB6 add-in TCP track**: no auth, loopback, only reachable when `VB6MCP_TRANSPORT` leaves it on.
- **mjs shell**: stdio, no network face.

## 1. Assets

| # | asset | why it matters |
|---|---|---|
| A-1 | the 120-tool face | it is arbitrary code execution *inside the user's VB6 IDE* and arbitrary file write in their project tree - `write_file`, `save_project`, `make_project`, `control_execution`, `quit_ide`, `add_component`, `git` |
| A-2 | the user's source tree and build outputs | read AND write through the same face |
| A-3 | the running IDE session | can be driven, broken, or closed; unsaved work can be lost |
| A-4 | the instance registry `instances\*.json` | pipe names, nonces, project paths, DLL sha - enough to attempt impersonation |
| A-5 | the MCP lesson store `mcp_lessons/` | global control-plane state, not secrets, but integrity-sensitive |

## 2. Adversaries (who we defend against)

### T1 - another user on the SAME machine  **== primary threat ==

**Reach**: loopback TCP is machine-wide, not user-scoped.  Any local account (or any service
running as another user) that can open a TCP connection to `127.0.0.1:<port>` gets the *whole* face
today.  They do not need to be an administrator and they do not need to touch the pipe.
**Consequence**: full A-1/A-2/A-3 - they can read the victim's sources, rewrite them, run code in
the victim's IDE, or close it.  **This is the gap that must close.**
**Capability note (honest)**: an unprivileged local user can already read the registry files as long
as they are ACL-readable, so T1 is not only about the TCP port; the port is what turns "read a
nonce" into "drive the IDE".

### T2 - a peer on the LAN

**Reach**: only if the bridge is bound to a non-loopback address, or a helper forwards for it
(`relay.exe`/portproxy/SSH tunnel).  `--host 0.0.0.0` was used twice experimentally on t16.
**Consequence**: identical to T1 but from off-box, i.e. the ADR's "the real hole".
**Position**: this threat is answered **structurally, not by the token**: see M-0.

### T3 - a compromised or malicious MCP client  **== OUT OF SCOPE for auth ==

By construction this party *holds the token*: it is the shell we hand it to.  A bearer token
authenticates **an endpoint**, never a user's intent, so it buys nothing here.  What would help is
per-tool authorization, an audit trail, and least privilege - none of which is this card.
**We must say this out loud** so that "token implemented" is never read as "the face is safe for
untrusted clients".

### T4 - a same-user local process

Same user = same rights.  It can read the token from the environment of a sibling process only with
debug privileges, but it can simply read a token FILE the user owns.  **Accepted**: defending a user
from themselves is not achievable, and pretending otherwise would be the dishonest option.
Mitigation is hygiene only: prefer an environment variable over a file, keep the file owner-only,
never put a token in the repository tree or in a committed config.

### T5 - accidental disclosure through our own artifacts

Tokens leak through logs, URLs (proxies, shell history, `access.log`), error messages and receipts.
This one is entirely on us and is **in scope** (M-3, M-4).

## 3. Reachability matrix

| surface | T1 same-machine | T2 LAN | T3 rogue client | T4 same user |
|---|---|---|---|---|
| bridge HTTP `/rpc`, `/mcp`, `/health` | **yes today** | only if non-loopback | yes (has token) | yes |
| add-in pipe | needs the nonce (registry) | n/a | n/a | yes |
| add-in TCP track | loopback only, when enabled | only if enabled + reachable | n/a | yes |
| mjs stdio | needs the process | no | n/a | yes |

## 4. Mechanism decision (derived above)

| # | decision | justified by |
|---|---|---|
| **M-0** | **Non-loopback binding is REMOVED, not merely defaulted against**: the bridge refuses any host that is not a loopback literal (`127.0.0.0/8`, `::1`, `localhost`).  Binding `0.0.0.0` becomes impossible rather than discouraged. | T2: a structural refusal cannot be forgotten by an operator at 2 a.m., and the ADR already forbids the exposure |
| **M-1** | **Every HTTP surface requires the token, including `/health`.**  Unauthenticated callers get 401 with a structured reason and **zero** body detail (no instances, no paths, no version). | T1: `/health` leaks the registry, project paths and the tool face today |
| **M-2** | Credential presented as `Authorization: Bearer <token>` **or** `X-VB6MCP-Token`; compared with a constant-time digest comparison (`hmac.compare_digest` over SHA-256 of both sides, so neither content nor length leaks through timing). | T1/T5: no early-exit comparison, no length oracle |
| **M-3** | The token is **never accepted in a URL**.  A query parameter or path segment carrying it is REJECTED with `token_in_url_rejected` (not silently honoured). | T5: URLs are logged by everything in the path |
| **M-4** | The token is **never echoed** and **never logged**.  Diagnostics (HTTP side) record only **length + source**, e.g. `auth_state enabled=true source=env tokenLen=17` - deliberately NOT even a fingerprint of the secret, because a published fingerprint is a guess-confirmation oracle.  A `redact()` guard is applied to every field the logger writes, so an accidental `token=...` cannot reach the file at all. | T5, and it is what makes the card's "logs/URL contain no token" assertion checkable |
| **M-5** | Source precedence: `VB6MCP_TOKEN` (environment, preferred - never touches disk) then `VB6MCP_TOKEN_FILE` (owner-only file).  If neither is set the service still runs, but `/health` reports `auth:"disabled"` and startup prints a warning. | T4 is accepted, so the weaker source is still allowed - but it must be visible |
| **M-6** | **Pipe face**: the bridge-side CLIENT carries the credential in the `hello` frame (`auth` object, section 5).  Server-side validation is a **separate card in the VB write scope**. | the card's item 3, and the ADR's reserved handshake slot |
| **M-7** | No rotation, no TTL, no per-tool policy in this card.  The field names are chosen so that a later `engineToken` (ADR sec 5.2) can be added without a wire break. | ADR sec 5.2 requires a renewed threat-model review before nonce persistence or token rotation |

## 5. Wire fields (pipe), to be pinned now and implemented server-side later

Additive only; a server that does not read them must be unaffected ('s parser looks up
`vb6mcp` and `nonce` only, so unknown keys are expected to be inert - **to be confirmed by **, and the
bridge will not depend on it: if a server answers `welcome` the token was not required).

| frame | field | type | meaning |
|---|---|---|---|
| `hello` | `auth` | object | `{"scheme":"bearer","token":"<secret>","tokenFp":"<sha256[:8]>"}` - `tokenFp` exists so a server can compare/audit without handling the secret in logs |
| `welcome` | `auth` | object | `{"state":"ok"}` or `{"state":"not_configured"}` |
| `deny` | `reason` | string | adds `token_mismatch` / `token_missing` to the existing set |

## 6. Non-goals (explicit)

* no defence against T3 (a client that legitimately holds the token);
* no transport encryption - loopback traffic is not a network-exposed channel, and TLS is out of scope;
* no per-tool authorization, no rate limiting, no user identity;
* no protection for the VB add-in's own TCP track beyond the binding rule (VB write scope);
* no secret store, no OS credential vault, no token rotation.

## 7. Verification plan (maps to the card's item 4)

| # | assertion | how it is proved |
|---|---|---|
| V1 | no token -> **401 `token_missing`** AND the tool did NOT run | the stub add-in's served-request counter is unchanged after the call |
| V2 | wrong token -> **401 `token_mismatch`** AND the tool did NOT run | same counter |
| V3 | token in URL -> **401 `token_in_url_rejected`** | request with `?token=...` is refused, never honoured |
| V4 | correct token -> behaviour is IDENTICAL to today | the existing selftest cases re-run with a token configured; tool face still 110/120, `schema_changes=0` |
| V5 | the log contains no token | assert the literal secret does not appear in the log file; only the fingerprint does |
| V6 | no receipt/URL echoes the token | assert the 401 body contains neither the provided value nor the fingerprint of the expected one |
| V7 | non-loopback binding is refused | `--host 0.0.0.0` (and a LAN IP) exit non-zero with a stated reason; loopback forms are accepted |
| V8 | write tool leaves no trace without a token | `tools/call` of a write tool with a bad token: counter unchanged, and no side-effect artifact appears |

## 8. Contract impact (reported to the Lead before implementation)

**Behaviour change, tool face unchanged.**  `tools/list` still returns the same 110/120 names with the
same `inputSchema`s; no new tool, no removed tool, no parameter change, so the five frozen artifacts
and `schema_changes=0` are unaffected.  What changes is the HTTP contract: an unauthenticated
request that used to succeed now gets 401 - which is the point of the card.  Two consequences for
other owners:

1. the **mjs thin shell** must send the token on `/health` and `/rpc` (a few lines in the
   planned bridge client); if it does not, it degrades to the snapshot face - visible, not silent.
2. any tooling that calls the bridge directly (the repo's audit/dump/smoke scripts) keeps working
   **when run with the token**, and gets an honest 401 when run without it.

The Lead decides batching and whether the contract version moves; this card does not touch the
version.

## 9. Residual risk after this card (to stay on the record)

* T1 is reduced to "an attacker who has read the token"; on a shared box an unprivileged user who
  can read an owner-only token file is still game over, and if the operator exports `VB6MCP_TOKEN` in a
  world-readable way the protection evaporates.  M-5 makes the weaker case visible, it does not
  prevent it.
* the pipe face is still nonce-only until the VB-side card lands: a same-machine attacker who can
  read the registry can still attempt impersonation (the nonce is the only proof today).
* T3 is untouched by design.


## 10. Remote mode (2026-10-17) - the retired invariant and its replacement

**RETIRED**: the old sentence "non-loopback binding is refused with exit 6 and there is deliberately
NO switch to override it".  It was true while the bridge had no TLS and no remote read-only tier;
keeping it now would either block the author-approved remote shape or turn a real capability into a
false statement.  What replaces it:

| gate | rule | refusal |
|---|---|---|
| default tier | loopback only; a non-loopback bind exits with **6** exactly as before | exit 6 |
| remote prerequisites | `--allow-remote <cidr[,cidr]>` + `--tls-cert/--tls-key` + a token source + a WRITABLE `--audit-log`; ALL of them | exit **7**, naming every missing item |
| source | the peer address is re-checked inside the process against the allowlist | `403 source_not_allowed` |
| privilege | a non-loopback caller is **READ-ONLY**: only an explicit read-only allowlist runs, everything else (including any NEW tool) is refused | `remote_read_only` unless `--remote-write` (loud at startup) |
| credential | constant-time compare, **two** valid tokens (current + previous) for rotation, per-source exponential backoff then lockout | `401 token_mismatch` / `429 auth_locked_out` |
| transport | TLS on the HTTP face only (TLS >= 1.2); a self-signed certificate is acceptable only when the client PINS the sha256 of the DER certificate | startup refuses exit 7 if TLS cannot load |
| audit | one append-only JSONL line per response (refusals included), never a token value; an unwritable log refuses remote requests | `503 audit_unavailable` |
| discovery | a remote caller never sees instance paths, pipe names, pids or the registry dir | `/health` redacted (`remoteRedacted:true`) |

**Why the gates are not optional**: a static bearer token is replayable, and this face executes code
in the IDE - so TLS and source restriction are the floor, not the ceiling.  A deny-first tool allowlist
makes a forgotten entry fail CLOSED instead of quietly shipping the write face to the network, and the
audit log is what makes an incident attributable at all.
