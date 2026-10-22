# Security

Multipanel is an MVP. It includes application-layer protections, but no software can promise perfect safety.

## Reporting a vulnerability

Please do not post working exploits, credentials, private prompts, or user data publicly. Open a private security report through the repository host, including:

- affected endpoint or feature
- reproducible steps without secrets
- impact and suggested mitigation
- whether the issue is already exposed in a public deployment

Rotate any exposed credential immediately and revoke it at the provider.

## Deployment requirements

Before a public launch:

- Store `GEMINI_API_KEY` only in the deployment secret manager.
- Use separate development and production keys and rotate them periodically.
- Set exact `ALLOWED_ORIGINS`; never use a wildcard.
- Put the API behind a WAF and distributed rate limiter.
- Keep databases and internal services private; this project currently has no database.
- Treat the Gemini API as an intentional external dependency. Private-subnet isolation cannot prevent that outbound model call without a private model endpoint.
- Enable logging and alerts for authentication/configuration failures, 4xx/5xx spikes, rate limits, and provider spend.
- Test the key-revocation procedure before launch.

See `README.md` for the implementation baseline and release checklist.
