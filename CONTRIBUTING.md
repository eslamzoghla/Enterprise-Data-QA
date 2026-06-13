# Contributing Guide

Thank you for interest in contributing to the Enterprise QA Engine!

## 🚀 Getting Started

### 1. Fork & Clone
```bash
git clone https://github.com/YOUR_USERNAME/Enterprise-Data-QA.git
cd Enterprise-Data-QA
git remote add upstream https://github.com/eslamzoghla/Enterprise-Data-QA.git
```

### 2. Create Feature Branch
```bash
git checkout -b feature/your-feature-name
```

### 3. Install Dependencies
```bash
npm install
```

### 4. Run Tests
```bash
npm test
```

---

## 📋 Development Guidelines

### Code Style
- **TypeScript** - Strict mode enabled
- **Naming:** camelCase for functions/variables, PascalCase for types/classes
- **Comments:** JSDoc for public functions
- **Line Length:** 100 characters max

### Module Development
1. Create file in `src/modules/`
2. Export clear interfaces
3. Add unit tests in `__tests__/` folder
4. Document in `ARCHITECTURE.md`

### Example: New Module
```typescript
// src/modules/myNewModule.ts
/**
 * Module X: My New Module
 * Description of what this module does
 */

export interface MyInput {
  // input structure
}

export interface MyOutput {
  // output structure
}

/**
 * Main function
 * @param input - The input data
 * @returns The output result
 */
export function myMainFunction(input: MyInput): MyOutput {
  // implementation
  return { /* ... */ }
}
```

---

## 🧪 Testing

### Write Tests
```typescript
// __tests__/myNewModule.test.ts
import { myMainFunction } from "../modules/myNewModule";

describe("myNewModule", () => {
  it("should do something", () => {
    const input = { /* ... */ };
    const output = myMainFunction(input);
    expect(output).toEqual({ /* ... */ });
  });
});
```

### Run Tests
```bash
npm test                    # All tests
npm test myNewModule        # Specific module
npm test -- --coverage      # With coverage report
```

---

## 📝 Commit Messages

Follow conventional commits:
```
feat: Add new error classification
fix: Correct Arabic normalization
docs: Update module documentation
test: Add test cases for suppression engine
refactor: Improve performance of alignment recovery
```

---

## 🔄 Pull Request Process

1. **Sync with upstream:**
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. **Push to your fork:**
   ```bash
   git push origin feature/your-feature-name
   ```

3. **Create Pull Request** on GitHub
   - Clear title and description
   - Link any related issues: "Closes #123"
   - List changes made
   - Include test results

4. **Respond to feedback** - Maintainers will review

5. **Merge** - Maintainer merges when approved

---

## 🐛 Reporting Issues

### Bug Report Template
```markdown
## Description
Clear description of the bug

## Steps to Reproduce
1. Step 1
2. Step 2

## Expected Behavior
What should happen

## Actual Behavior
What actually happens

## Screenshots
[If applicable]

## Environment
- OS: [e.g., macOS]
- Node version: [e.g., 18.0.0]
- Browser: [e.g., Chrome]
```

### Feature Request Template
```markdown
## Description
Clear description of the feature

## Use Case
Why this feature is needed

## Proposed Solution
How it should work

## Alternatives
Other solutions considered
```

---

## 📚 Areas for Contribution

### High Priority
- [ ] PDF export implementation
- [ ] Batch processing for 1000+ audits
- [ ] Performance optimizations
- [ ] API endpoints for external integrations

### Medium Priority
- [ ] Multi-language support (Spanish, French)
- [ ] Fuzzy table matching
- [ ] Real-time dashboards
- [ ] Advanced filtering UI

### Low Priority
- [ ] Dark mode theme
- [ ] Export format options
- [ ] Historical trend analysis
- [ ] User preference storage

---

## 💡 Development Tips

### Debug Module Execution
```typescript
import { executeQAEvaluation } from "./utils/qaEngineOrchestrator";

// Enable verbose logging
console.log = (...args) => originalLog("[QA]", ...args);

const result = await executeQAEvaluation(emp, rev, config);
console.log(result); // See detailed output
```

### Test with Sample Data
```typescript
import { getDemoEmployeeData, getDemoReviewerData } from "./utils/demoData";

const empWb = getDemoEmployeeData();
const revWb = getDemoReviewerData();
const config = { /* test config */ };
```

### Profile Performance
```typescript
const start = performance.now();
const result = await executeQAEvaluation(empWb, revWb, config);
const duration = performance.now() - start;
console.log(`Total time: ${duration.toFixed(2)}ms`);
```

---

## 📖 Resources

- [TypeScript Documentation](https://www.typescriptlang.org/docs/)
- [React Best Practices](https://react.dev/learn)
- [Excel.js Library](https://github.com/exceljs/exceljs)
- [Arabic NLP](https://en.wikipedia.org/wiki/Arabic_language_processing)
- [Levenshtein Distance](https://en.wikipedia.org/wiki/Levenshtein_distance)

---

## ❓ Questions?

- Open a [GitHub Discussion](https://github.com/eslamzoghla/Enterprise-Data-QA/discussions)
- Check existing [Issues](https://github.com/eslamzoghla/Enterprise-Data-QA/issues)
- Email: contact@example.com

---

## 🎉 Thank You!

Your contributions make this project better. We appreciate:
- Bug reports
- Feature suggestions
- Code contributions
- Documentation improvements
- Spreading the word!

**Happy coding! 🚀**
