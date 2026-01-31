# Contributing to EmmaNeigh

Thank you for your interest in improving the Signature Packet Builder! This document provides guidelines for contributing to the project.

---

## Table of Contents

1. [Code of Conduct](#code-of-conduct)
2. [How to Contribute](#how-to-contribute)
3. [Development Setup](#development-setup)
4. [Coding Standards](#coding-standards)
5. [Testing Guidelines](#testing-guidelines)
6. [Pull Request Process](#pull-request-process)
7. [Feature Requests](#feature-requests)
8. [Bug Reports](#bug-reports)

---

## Code of Conduct

### Our Standards

This is an internal legal tool. Contributors should:

- Be professional and respectful
- Focus on what's best for users (attorneys and staff)
- Accept constructive criticism gracefully
- Prioritize security and confidentiality
- Assume good faith

### Scope

This code applies to:
- GitHub interactions (issues, PRs, comments)
- Code reviews
- Email communications about the project

---

## How to Contribute

### Types of Contributions

We welcome:

1. **Bug fixes** - Especially for signer detection edge cases
2. **Feature enhancements** - See [Roadmap](ARCHITECTURE.md#future-roadmap)
3. **Documentation improvements** - Clarity, examples, FAQs
4. **Testing** - New test cases, performance benchmarks
5. **User experience** - Error messages, logging, output formatting

### Not Accepting

We are **not** currently accepting:
- GUI rewrites (intentionally console-based)
- Cloud integrations (must remain local-only)
- External API dependencies
- Anything that requires user installation

---

## Development Setup

### Prerequisites

- Python 3.11+ installed on development machine
- Git installed
- Text editor or IDE (VS Code recommended)
- Sample transaction PDFs for testing

### Initial Setup

```bash
# 1. Fork the repository on GitHub

# 2. Clone your fork
git clone https://github.com/YOUR_USERNAME/EmmaNeigh.git
cd EmmaNeigh

# 3. Create development branch
git checkout -b feature/your-feature-name

# 4. Install dependencies (if modifying code)
pip install pymupdf pandas openpyxl

# 5. Test current version
python src/build_signature_packets.py path/to/test/pdfs
```

### Project Structure

```
EmmaNeigh/
├── src/
│   └── build_signature_packets.py   # Core logic - this is what you'll modify
├── python/                           # Portable runtime (don't modify)
├── docs/                             # Documentation
├── run_signature_packets.bat         # Launcher
└── README.md                         # Main documentation
```

---

## Coding Standards

### Python Style

**Follow PEP 8** with these specifics:

- **Line length:** 88 characters (Black formatter default)
- **Indentation:** 4 spaces
- **Quotes:** Double quotes for strings
- **Naming:**
  - Functions: `snake_case`
  - Classes: `PascalCase`
  - Constants: `UPPER_CASE`

### Documentation

**All functions must have docstrings:**

```python
def extract_person_signers(text):
    """
    Extract individual person names from a signature page.
    
    Args:
        text (str): Full text content of a PDF page
        
    Returns:
        set: Set of normalized person names found on the page
        
    Example:
        >>> text = "BY: ___\\nName: John Smith\\nTitle: President"
        >>> extract_person_signers(text)
        {'JOHN SMITH'}
    """
    # Implementation...
```

### Comments

**Use comments for:**
- Non-obvious logic
- Algorithm explanations
- Security-critical sections
- Workarounds for library quirks

**Example:**
```python
# Look upward from BY: marker to find entity name
# This is Tier 2 fallback when explicit Name: field is missing
for j in range(1, 7):
    candidate = lines[by_index - j]
    # ...
```

### Error Messages

**User-facing errors must be clear:**

❌ Bad:
```python
raise Exception("Error in function X")
```

✅ Good:
```python
raise RuntimeError(
    "No signers detected in PDFs. "
    "Ensure documents have standard signature blocks with 'BY:' and 'Name:' fields."
)
```

---

## Testing Guidelines

### What to Test

1. **Standard signature blocks** (credit agreements, guaranties)
2. **Edge cases:**
   - Multiple signers per page
   - Same person in multiple documents
   - Unusual name formats (hyphenated, suffixes)
3. **Error conditions:**
   - Empty folder
   - No signature pages found
   - Malformed PDFs
4. **Performance:**
   - Large deals (100+ documents)
   - Many signers (100+)

### Test Data

**Create sanitized PDFs:**
- Remove real client names
- Use placeholder entities ("ABC Corp", "John Smith")
- Preserve signature block structure

**Example Test Cases:**

```
test_pdfs/
├── standard_signature_block.pdf       # Typical credit agreement
├── multiple_signers_one_page.pdf      # Joinder with 5 signers
├── same_person_three_docs.pdf         # John Smith signs 3 times
├── no_name_field.pdf                  # BY: without explicit Name:
└── scanned_without_ocr.pdf            # Should fail gracefully
```

### Running Tests

```bash
# Manual testing
python src/build_signature_packets.py test_pdfs/

# Verify output
ls test_pdfs/signature_packets_output/packets/
```

### Regression Testing

Before submitting PR:
1. Run tool on your test PDFs
2. Compare output to previous version
3. Document any changes in behavior

---

## Pull Request Process

### Before Submitting

**Checklist:**
- [ ] Code follows style guidelines
- [ ] All functions have docstrings
- [ ] No hardcoded paths or credentials
- [ ] Tested on sample PDFs
- [ ] Documentation updated
- [ ] No new external dependencies

### PR Title Format

Use clear, descriptive titles:

✅ Good:
- `Fix: Handle signature blocks without explicit Name: field`
- `Feature: Add support for initial-only pages`
- `Docs: Clarify installation instructions for Windows`

❌ Bad:
- `Update code`
- `Fix bug`
- `Changes`

### PR Description Template

```markdown
## Description
Brief explanation of changes

## Motivation
Why is this change needed?

## Testing
How was this tested?

## Screenshots (if applicable)
Before/after outputs

## Checklist
- [ ] Code tested
- [ ] Documentation updated
- [ ] No breaking changes
```

### Review Process

1. **Automated checks** (if configured)
2. **Maintainer review** (code quality, security)
3. **Testing on real transaction PDFs**
4. **Approval and merge**

### After Merge

- Your changes will be included in the next release
- You'll be credited in the changelog

---

## Feature Requests

### How to Request

1. **Search existing issues** - May already be planned
2. **Open new issue** with label `enhancement`
3. **Describe use case:**
   - What problem does this solve?
   - How common is this scenario?
   - Proposed implementation (if you have ideas)

### Feature Evaluation Criteria

We consider:
- **User impact:** How many people benefit?
- **Complexity:** Implementation effort vs. value
- **Security:** Does it maintain local-only processing?
- **Maintenance:** Can we support it long-term?

### Roadmap Priorities

Current focus areas:
1. Initial-only page detection
2. Better duplicate name handling
3. Capacity field extraction (Borrower/Guarantor)
4. Cover sheet generation

---

## Bug Reports

### How to Report

**Use GitHub Issues** with label `bug`

**Include:**
1. **Description:** What happened vs. what you expected
2. **Steps to Reproduce:**
   ```
   1. Run tool on folder with X documents
   2. Observe output
   3. See error...
   ```
3. **Environment:**
   - OS: Windows 10, macOS Monterey, etc.
   - Python version: `python --version`
   - Tool version: v1.0.0
4. **Logs:** Copy console output
5. **Sample PDF (if possible):**
   - Sanitize client data
   - Recreate issue with fake data

### Bug Priority Levels

**Critical:**
- Tool crashes
- Data loss
- Security vulnerability

**High:**
- Incorrect signature packet generation
- Missing pages for signers

**Medium:**
- Poor error messages
- Performance issues

**Low:**
- Cosmetic issues
- Nice-to-have features

---

## Development Workflow

### Branching Strategy

```
main              # Production-ready code
  ├── develop     # Integration branch (if needed)
  └── feature/*   # Individual features
      └── bugfix/*   # Bug fixes
```

### Commit Messages

Follow conventional commits:

```
<type>: <description>

[optional body]

[optional footer]
```

**Types:**
- `feat:` New feature
- `fix:` Bug fix
- `docs:` Documentation only
- `refactor:` Code restructuring
- `test:` Adding tests
- `chore:` Maintenance

**Examples:**
```
feat: add initial-only page detection

fix: handle signature blocks without Name: field
Closes #42

docs: clarify Windows installation steps
```

---

## Code Review Guidelines

### For Reviewers

**Focus on:**
- Correctness (does it work?)
- Security (any new risks?)
- Maintainability (can others understand it?)
- Documentation (is it explained?)

**Be constructive:**
- Explain *why* changes are needed
- Suggest alternatives
- Acknowledge good work

### For Contributors

**Responding to feedback:**
- Assume good intent
- Ask clarifying questions
- Make requested changes or discuss alternatives
- Update PR description if scope changes

---

## Release Process

*(For maintainers)*

### Version Numbering

Semantic Versioning: `MAJOR.MINOR.PATCH`

- **MAJOR:** Breaking changes
- **MINOR:** New features (backward compatible)
- **PATCH:** Bug fixes

### Release Checklist

1. Update `CHANGELOG.md`
2. Tag release: `git tag v1.1.0`
3. Build distribution ZIP
4. Create GitHub release
5. Update documentation if needed

---

## Questions?

**Not sure where to start?**
- Check [Issues labeled "good first issue"](https://github.com/raamtambe/EmmaNeigh/labels/good%20first%20issue)
- Ask questions in issue comments
- Email project maintainer

**Have an idea but don't code?**
- Open a feature request
- Describe your use case
- We'll find someone to implement it

---

## Recognition

Contributors will be:
- Listed in `CONTRIBUTORS.md`
- Credited in release notes
- Thanked publicly (if desired)

---

## License

By contributing, you agree that your contributions will be licensed under the same MIT License that covers the project.

---

**Thank you for helping make EmmaNeigh better!**
