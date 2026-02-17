# Contributing to D365 Copilot Toolbox

Thank you for your interest in contributing! We'd **love** your help making D365 multi-agent workflows better.

## Ways to Contribute

- **New Solutions** — create new agent integrations, workflows, or examples
- **Features** — implement new functionality in existing solutions
- **Bug Fixes** — find and fix issues
- **Documentation** — improve guides, add examples, fix typos
- **Testing** — report bugs, verify fixes
- **Ideas** — suggest new features or solutions via issues or email

## Getting Started

### Development Environment

1. **D365 F&O development VM** (version 10.0.45+)
2. **Visual Studio** with D365 development tools
3. **Git** for version control

### Local Setup

1. Fork and clone the repository:
   ```bash
   git clone https://github.com/<your-fork>/D365CopilotToolbox.git
   ```

2. Register symbolic links (run as Administrator):
   ```powershell
   cd D365CopilotToolbox\Scripts
   .\RegisterSymbolicLinks.ps1
   ```

3. Download vendor JavaScript libraries (requires Node.js/npm on PATH):
   ```powershell
   .\Update-VendorLibs.ps1
   ```
   This reads `Scripts/vendor-libs.json` and downloads the required npm packages (MSAL, WebChat, Copilot Studio Client) into the AxResource content folders. Use `-Force` to re-download if files already exist.

4. Open the solution in Visual Studio:
   - `Project\CopilotAgentHost\CopilotAgentHost.sln`

5. Build and synchronize the database

## Coding Standards

### Naming Conventions

| Object Type | Prefix | Example |
|-------------|--------|---------|
| All objects in Copilot Toolbox | `COTX` | `COTXCopilotHostControl` |
| Example model objects | `CTXE` | `CTXESalesTable` |
| Labels | `COTX` | `COTXAgentParameters` |

### CSS Standards

- Scope all selectors under `[data-dyn-role="COTXCopilotHostControl"]`
- Use px units for predictable sizing within the D365 form chrome

## Pull Request Process

1. **Create a feature branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes** following the coding standards above

3. **Test locally:**
   - Build succeeds with no errors
   - BP checks pass (run Best Practices analyzer in Visual Studio)
   - Functional testing on a development environment

4. **Commit with descriptive messages:**
   ```
   feat: add support for custom context providers
   fix: resolve label BP error on agent parameters table
   docs: add troubleshooting section to getting-started guide
   ```

5. **Push and open a Pull Request** against `main`

6. **PR description should include:**
   - What the change does
   - Why it's needed
   - How to test it
   - Screenshots (for UI changes)

## Reporting Issues

1. Search [existing issues](../../issues) first
2. If not found, create a new issue with:
   - D365 version
   - Steps to reproduce
   - Expected vs actual behavior
   - Error messages, stack traces, or screenshots

## Feature Requests

Open a [GitHub issue](../../issues/new) with the `enhancement` label, or email **copilot@erpilots.com**.

## Code of Conduct

- Be respectful and constructive
- Focus on the code, not the person
- Help newcomers learn the codebase
- Give credit where it's due

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
