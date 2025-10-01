onderhold: is there something like a "build" command in Python?

GitHub Copilot: Yes, in Python packaging, a "build" command is commonly used to create distribution files for your project.

**Standard build command:**
```sh
python -m build
```
This uses the [`build`](https://pypi.org/project/build/) package to generate source (`.tar.gz`) and wheel (`.whl`) distributions in the dist directory.

**Steps:**
1. Install the build tool (if not already):
   ```sh
   pip install build
   ```
2. Run the build command:
   ```sh
   python -m build
   ```

**Summary:**  
- `python -m build` is the standard way to build Python packages.
- It reads your pyproject.toml and creates distributable files.

onderhold: how and where do offer binaries in my github repo, and will building those also included in ````python -m build````?

GitHub Copilot: ### Offering binaries in your GitHub repo

1. **Release Assets:**  
   The most common way is to upload your built binaries (e.g., `.whl`, `.tar.gz`, `.exe`) as assets to a [GitHub Release](https://docs.github.com/en/repositories/releasing-projects-on-github/creating-releases).
   - Go to your repo’s "Releases" page.
   - Create a new release (or edit an existing one).
   - Attach your binary files.

2. **CI/CD Automation:**  
   You can automate building and uploading binaries using GitHub Actions.  
   Example: After building with `python -m build`, use an action to upload the files to the release.

---

### Does `python -m build` include binaries?

- `python -m build` creates **Python package distributions**:  
  - Source distribution (`.tar.gz`)
  - Wheel distribution (`.whl`)
- These are placed in the dist folder.
- **It does NOT create standalone executables** (like `.exe`).  
  For that, use tools like `pyinstaller` or `cx_Freeze`.

**Summary:**  
- Use `python -m build` to create `.whl` and `.tar.gz` for PyPI and releases.
- Upload these files as release assets on GitHub.
- For true binaries (`.exe`), use a packaging tool and upload those separately.

onderhold: can I also store the binary executables somewhere in my local repo, and then push it to Github in a way that they will be included in a release?

GitHub Copilot: Yes, you can store binary executables (e.g., `.exe`, `.whl`, `.tar.gz`) in your local repo, but **it is not recommended to commit them directly to your repository** because:

- They will bloat your repository size.
- They are not source code and are better managed as release assets.

**Recommended workflow:**

1. **Build your binaries locally** (e.g., with `python -m build` or `pyinstaller`).
2. **Do NOT commit them to your repo.**
3. Instead, when you want to publish a release:
   - Go to your repo’s "Releases" page on GitHub.
   - Create a new release (or edit an existing one).
   - Upload your binaries as release assets.

**Automated alternative:**  
Use GitHub Actions to build your binaries and automatically attach them to a release.

**Summary:**  
- Build binaries locally.
- Upload them as release assets on GitHub.
- Do not commit binaries to your source repository.  
This keeps your repo clean and follows best practices.

onderhold: is there some sort of tool that I can use to add an .exe to a Github release from a local directory which is not in the local repo?

GitHub Copilot: Yes, you can use the official GitHub CLI tool [`gh`](https://cli.github.com/) to upload an `.exe` (or any file) from a local directory to a GitHub release, **even if the file is not in your local repo**.

### Example workflow

1. **Install GitHub CLI:**  
   [Download and install](https://cli.github.com/) for your platform.

2. **Authenticate:**  
   ```sh
   gh auth login
   ```

3. **Create a release (if needed):**  
   ```sh
   gh release create v1.0.0 --title "Release v1.0.0" --notes "Release notes here"
   ```

4. **Upload your `.exe` to the release:**  
   ```sh
   gh release upload v1.0.0 path/to/your/file.exe
   ```

You can upload any file from any location on your disk—**it does not need to be tracked by git or in your repo folder**.

**Summary:**  
Use `gh release upload <tag> <file>` to add binaries to a GitHub release from anywhere on your system.