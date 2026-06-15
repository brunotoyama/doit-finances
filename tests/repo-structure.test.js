import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, '..');

describe('Repository Structure Validation', () => {
  /**
   * Validates: Requirements 2.1, 2.6
   * All required files must exist at the project root.
   */
  describe('Required files exist', () => {
    const requiredFiles = [
      'index.html',
      'script.js',
      'style.css',
      'README.md',
      'LICENSE',
      '.gitignore',
      'package.json'
    ];

    requiredFiles.forEach((file) => {
      it(`should have ${file}`, () => {
        const filePath = path.join(ROOT, file);
        expect(fs.existsSync(filePath)).toBe(true);
      });
    });
  });

  /**
   * Validates: Requirements 2.4, 2.5
   * Forbidden files and directories must NOT exist in the published repo.
   * Note: node_modules is checked via .gitignore since it exists during development.
   */
  describe('Forbidden files do NOT exist', () => {
    const forbiddenPaths = [
      { name: '0', type: 'file' },
      { name: '.kiro', type: 'directory' },
      { name: '.DS_Store', type: 'file' }
    ];

    forbiddenPaths.forEach(({ name, type }) => {
      it(`should NOT have ${type} "${name}"`, () => {
        const targetPath = path.join(ROOT, name);
        expect(fs.existsSync(targetPath)).toBe(false);
      });
    });

    it('should have node_modules/ listed in .gitignore', () => {
      const gitignore = fs.readFileSync(path.join(ROOT, '.gitignore'), 'utf-8');
      expect(gitignore).toContain('node_modules');
    });
  });

  /**
   * Validates: Requirements 1.6, 3.1
   * package.json must be valid JSON with correct scripts configuration.
   */
  describe('package.json validation', () => {
    let pkg;

    it('should be valid JSON', () => {
      const content = fs.readFileSync(path.join(ROOT, 'package.json'), 'utf-8');
      expect(() => {
        pkg = JSON.parse(content);
      }).not.toThrow();
    });

    it('should have start script containing "serve"', () => {
      const content = fs.readFileSync(path.join(ROOT, 'package.json'), 'utf-8');
      pkg = JSON.parse(content);
      expect(pkg.scripts).toBeDefined();
      expect(pkg.scripts.start).toBeDefined();
      expect(pkg.scripts.start).toContain('serve');
    });

    it('should have start script containing "3000"', () => {
      const content = fs.readFileSync(path.join(ROOT, 'package.json'), 'utf-8');
      pkg = JSON.parse(content);
      expect(pkg.scripts.start).toContain('3000');
    });

    it('should NOT have a "dependencies" field with items (only devDependencies allowed)', () => {
      const content = fs.readFileSync(path.join(ROOT, 'package.json'), 'utf-8');
      pkg = JSON.parse(content);
      if (pkg.dependencies) {
        expect(Object.keys(pkg.dependencies).length).toBe(0);
      }
    });
  });

  /**
   * Validates: Requirements 1.1, 1.2, 2.5
   * No file in the repo should reference "kiro" or "streamlit".
   */
  describe('No references to kiro or streamlit', () => {
    const filesToCheck = [
      'index.html',
      'script.js',
      'style.css',
      'README.md',
      'package.json'
    ];

    filesToCheck.forEach((file) => {
      it(`${file} should not contain references to "kiro"`, () => {
        const filePath = path.join(ROOT, file);
        if (fs.existsSync(filePath)) {
          const content = fs.readFileSync(filePath, 'utf-8').toLowerCase();
          expect(content).not.toContain('kiro');
        }
      });

      it(`${file} should not contain references to "streamlit"`, () => {
        const filePath = path.join(ROOT, file);
        if (fs.existsSync(filePath)) {
          const content = fs.readFileSync(filePath, 'utf-8').toLowerCase();
          expect(content).not.toContain('streamlit');
        }
      });
    });
  });
});
