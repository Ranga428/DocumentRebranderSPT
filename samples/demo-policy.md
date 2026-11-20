# Access Control Policy

This policy defines how Sapient staff and systems access production environments. It applies to every employee, contractor and automated agent.

## Scope

- All production and staging systems
- Source control and deployment tooling
- Client-facing APIs

## Requirements

1. Use named accounts; shared credentials are prohibited.
2. Enforce multi-factor authentication for administrative access.
3. Rotate service credentials every 90 days.
4. Review access grants each quarter.

## Responsibilities

| Role | Applies to | Behaviour |
| --- | --- | --- |
| System owner | All production systems | Approves access and reviews grants |
| Security team | All staff | Monitors authentication events and responds to alerts |
| Platform team | Deployment tooling | Maintains the identity provider configuration |

## Exceptions

Any exception must be documented, time-bound and approved by the Security Lead.

## Review

This policy is reviewed annually, or sooner after a material change to the environment.
