# Contributing to DataWizard

Thank you for your interest in contributing to DataWizard! This document provides guidelines and instructions for contributing.

## 🎯 How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the existing issues to avoid duplicates.

**When filing a bug report, include:**
- Clear, descriptive title
- Steps to reproduce the issue
- Expected vs. actual behavior
- Sample CSV/Excel files (if possible)
- DataWizard version
- .NET Framework version
- Operating system

**Example:**
```
Title: German number format not recognized in column X

Steps:
1. Load CSV with German numbers (123,45)
2. Call csv.WriteXLSX()
3. Open Excel file

Expected: Number displayed as 123,45
Actual: Number displayed as 12345

Sample file: attached test.csv
Version: DataWizard 1.0.0
```

### Suggesting Features

Feature suggestions are welcome! Please include:
- Clear use case
- Expected behavior
- Why this would benefit other users
- Any implementation ideas (optional)

### Pull Requests

1. **Fork** the repository
2. **Create a branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Make your changes** with clear commit messages
4. **Test** your changes thoroughly
5. **Update documentation** if needed
6. **Submit** a pull request

## 🔧 Development Setup

### Prerequisites

- Visual Studio 2017 or later
- .NET Framework 4.7.2 SDK
- Git

### Building the Project

```bash
# Clone your fork
git clone https://github.com/your-username/DataWizard.git
cd DataWizard

# Restore NuGet packages
nuget restore DataWizard.sln

# Build
msbuild DataWizard.sln /p:Configuration=Release
```

### Project Structure

```
DataWizard/
├── DataWizard/              # Main CSV library
│   ├── CSV.cs              # CSV parser and analyzer
│   └── DataWizard.csproj
├── libDataWizard/          # Excel library
│   ├── XLS.cs              # Excel operations
│   ├── DataWizardConfig.cs # Configuration
│   └── libDataWizard.csproj
├── test/                   # Test project
│   ├── Program.cs          # Examples
│   └── test.csproj
└── DataWizard.sln          # Solution file
```

## 📝 Coding Guidelines

### C# Style

- **Indentation**: 4 spaces (no tabs)
- **Braces**: K&R style (opening brace on same line)
- **Naming**:
  - Classes: `PascalCase`
  - Methods: `PascalCase`
  - Private fields: `camelCase` or `_camelCase`
  - Properties: `PascalCase`
  - Constants: `UPPER_CASE`

**Example:**
```csharp
public class CsvParser
{
    private char separator;
    private int _fieldCount;
    
    public bool IsValid { get; set; }
    
    public void ParseLine(string line)
    {
        // Implementation
    }
}
```

### Documentation

- **XML Comments** for public methods and classes
- **Inline comments** for complex logic
- **Update documentation** when changing behavior

**Example:**
```csharp
/// <summary>
/// Parses a CSV line and splits it into fields
/// </summary>
/// <param name="line">The CSV line to parse</param>
/// <param name="separator">The field separator character</param>
/// <returns>Array of parsed fields</returns>
public CsvField[] SplitLine(string line, char separator)
{
    // Implementation
}
```

### Error Handling

- **Use exceptions** for exceptional conditions
- **Validate inputs** at public API boundaries
- **Provide context** in exception messages
- **Clean up resources** properly (using, IDisposable)

**Example:**
```csharp
public void Load(string filePath)
{
    if (string.IsNullOrWhiteSpace(filePath))
        throw new ArgumentNullException(nameof(filePath));
        
    if (!File.Exists(filePath))
        throw new FileNotFoundException($"CSV file not found: {filePath}");
        
    using (StreamReader reader = new StreamReader(filePath))
    {
        // Process file
    }
}
```

## 🧪 Testing

### Manual Testing

Test your changes with:
- Various CSV formats (different separators, encodings)
- Edge cases (empty files, single column, special characters)
- Both directions (CSV→Excel and Excel→CSV)
- Different Excel sources (Excel, LibreOffice, Google Sheets)

### Test Cases to Cover

1. **CSV Parsing**
   - Different separators: `;`, `,`, `\t`, `|`
   - Quoted fields with escaped quotes
   - Empty fields
   - Multi-line fields

2. **Number Formats**
   - German format: `123,45`
   - English format: `123.45`
   - Large numbers: `1234567.89`
   - Small numbers: `0.001`

3. **Date Handling**
   - ISO format: `2025-01-31`
   - German format: `31.01.2025`
   - US format: `01/31/2025`
   - Edge case: numbers that look like dates

4. **Encoding**
   - UTF-8
   - UTF-8 with BOM
   - Windows-1252
   - Special characters: ä, ö, ü, ß, €

5. **Excel Features**
   - Multi-sheet workbooks
   - Different data types
   - Cell formatting
   - Formula cells

### Example Test

```csharp
// Test: German number format
CSV csv = new CSV();
csv.Load("test_german_numbers.csv");

// Verify separator detected
Assert.AreEqual(';', csv.Separator);

// Convert to Excel
csv.WriteXLSX("output.xlsx", true);

// Convert back to CSV
XLS.ToCsv("output.xlsx", "roundtrip.csv");

// Verify data integrity
// (Compare original and roundtrip files)
```

## 📚 Documentation

### When to Update Documentation

- New features → Update README.md
- API changes → Update XML comments
- Bug fixes → Update CHANGELOG.md
- Configuration → Update CONFIG_DOCUMENTATION.md

### Documentation Files

- **README.md**: Main documentation
- **CHANGELOG.md**: Version history
- **CONTRIBUTING.md**: This file
- **CONFIG_DOCUMENTATION.md**: Configuration guide
- **docs/**: Additional guides

## 🔄 Git Workflow

### Commit Messages

Follow the [Conventional Commits](https://www.conventionalcommits.org/) format:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation
- `style`: Code style (formatting)
- `refactor`: Code refactoring
- `test`: Tests
- `chore`: Maintenance

**Examples:**
```
feat(csv): Add support for tab separator

Add \t (tab) as a supported separator alongside ; , |

Closes #123

---

fix(excel): Prevent false date detection for numbers < 100

Numbers like 6 or 45.66 were incorrectly interpreted as dates.
Added heuristic: numbers < 100 are never treated as dates.

Fixes #456

---

docs(readme): Update installation instructions

Add dotnet build command and prerequisites section
```

### Branch Naming

- `feature/feature-name` - New features
- `fix/bug-description` - Bug fixes
- `docs/what-changed` - Documentation
- `refactor/what-changed` - Refactoring

### Pull Request Process

1. **Update** your branch with latest `main`
2. **Ensure** all tests pass
3. **Update** CHANGELOG.md with your changes
4. **Write** a clear PR description:
   - What changed
   - Why it changed
   - How to test it
5. **Request** review from maintainers
6. **Address** review comments
7. **Merge** after approval

## 🎨 Code Review

### What We Look For

- **Correctness**: Does it work as intended?
- **Performance**: Any performance implications?
- **Readability**: Is the code clear and maintainable?
- **Tests**: Are edge cases covered?
- **Documentation**: Is it documented?

### Review Checklist

- [ ] Code follows style guidelines
- [ ] XML comments for public APIs
- [ ] No breaking changes (or clearly documented)
- [ ] Tests added/updated
- [ ] Documentation updated
- [ ] CHANGELOG.md updated
- [ ] Commit messages are clear

## 🐛 Common Issues

### Issue: Cannot build project

**Solution:**
```bash
# Restore NuGet packages
nuget restore DataWizard.sln

# Clean and rebuild
msbuild DataWizard.sln /t:Clean
msbuild DataWizard.sln /t:Build
```

### Issue: Tests fail

**Solution:**
- Check file paths in test code
- Ensure test files exist in test/Samples/
- Verify encoding of test CSV files
- Check .NET Framework version

### Issue: Git conflicts

**Solution:**
```bash
# Update your branch
git fetch origin
git rebase origin/main

# Resolve conflicts
# ... edit files ...

git add .
git rebase --continue
```


## 📜 License

By contributing to DataWizard, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to DataWizard! 🎉
