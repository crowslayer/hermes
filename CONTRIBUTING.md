# Contributing to Hermes

Thank you for your interest in contributing to Hermes! This document provides guidelines and instructions for contributing.

## How to Contribute

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates.

When creating a bug report, include:

- **Clear title and description**
- **Steps to reproduce**
- **Expected behavior**
- **Actual behavior**
- **Environment details** (Excel version, Outlook version, Windows version)

### Suggesting Enhancements

Enhancement suggestions are welcome! Please provide:

- **Use case** - Why is this enhancement needed?
- **Proposed solution** - How should it work?
- **Alternatives considered** - What other approaches did you think about?

### Pull Requests

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Development Guidelines

### VBA Code Style

- Use **meaningful variable names** (camelCase for variables, PascalCase for functions)
- Add **comments** for complex logic
- Keep modules **focused** on single responsibility
- Handle errors gracefully with `On Error`

### Testing

Before submitting a PR:

1. Test with a small dataset (5-10 rows)
2. Verify preview mode works correctly
3. Verify auto-send mode works correctly
4. Check that attachments are matched correctly
5. Ensure status columns are updated properly

### Commit Messages

- Use present tense ("Add feature" not "Added feature")
- Use imperative mood ("Move cursor" not "Moves cursor")
- Reference issues when applicable ("Fixes #123")

## Code of Conduct

- Be respectful and inclusive
- Focus on constructive feedback
- Help newcomers feel welcome

## Questions?

Open an issue with the label "question" if you have any questions about contributing.
