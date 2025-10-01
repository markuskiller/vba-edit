## Possible Additional Features and Improvements

1. Startup logic: Check if the document is already open in another instance of the application.
2. Using the Office applications in invisible mode (currently, the application window is always shown).
3. Add CLI option to use a config file (e.g. `excel-vba.toml`).
   - The config file can be specified via a command-line option (e.g. `--config path/to/config.toml`).
   - There's no provision for a default config file, it always has to be specified explicitly.
   - The config file can contain options such as the input and output file paths, the Office application to use, and any additional command-line arguments.
   - Placeholders in the config file (e.g. `{file.name}`, `{file.path}`, `{file.fullname}`) will be resolved based on the actual file being processed.
   - The config file will support sections for different settings, such as `general`, `word`, `excel`, etc.
   - Command-line arguments will override config file settings.
4. exporting/importing PowerQuery M Code
5. replace watchgod with newer library `watchfiles`

- (very optional:)
    - implement a (readonly) virtual filesystem that mirrors the structure of the VBA project
    - exporting additional Office document properties (e.g. custom document properties, custom XML parts, references, etc.) or even the entire document (e.g. using `python-docx`, `openpyxl`, `python-pptx`, `pywin32` etc.)

- execution or mode flag: this utility may be used in two different use cases:
  1. as a one-off tool to export Office program code to the filesystem in a CI/CD pipeline or similar
  2. as a short running process to import some code modules once (mix-in)
  3. as a long-running process that keeps the file system and the VBA editor in sync (edit mode)
  
  The current implementation tries to cover both use cases, but it might be cleaner to have separate commands or modes for these two scenarios.
  - as a one-off exporting tool, it might be useful to have a command that just opens the document in readonly mode using a separate invisible instance of the Office application (irrespective of whether the document is already open in another instance), exports the code, and then ends the application again (closing the document implicitly as well).
  - when used to add some code modules once (mix-in), the document shall be opened visibly. There will be the need for the user to recompile, so the application needs to be visible and stay opened.
  - since edit mode always means that changes to the code can potentially happen concurrently in two places (VBA editor and file system), there needs to be a clear strategy for resolving conflicts and keeping both sides in sync. This is the most complex use case and might benefit from a more robust implementation, e.g. using a virtual filesystem. But even with a virtual filesystem, there's no way to control in-memory changes at neither the single module file level nor in the VBA editor. So, a last writer-wins strategy is the only feasible approach. At least, there should be a mechanism to detect conflicts of saved content and wether unsaved changes exist, and notify the user accordingly (e.g. using hash codes for each module for saved content, and just warning messages about unsaved changes).

### Configuration File

- The configuration file is a TOML file that allows users to customize the behavior of the CLI tool. Design decisions include:
  - File format: TOML is chosen for its simplicity and readability. It looks like INI files but supports nested structures.
  - Location: The config file is specified via a command-line option; if a path is omitted, it defaults to the current working directory.
  - Structure: The config file will have sections for different settings, such as default file paths, encoding options, and application-specific settings.

- Users can specify options such as the input and output file paths, the Office application to use, and any additional command-line arguments.

### Invisible Mode

- The invisible mode allows users to run the Office application in the background without displaying the application window. This mode will spawn a new instance of the application.
- Before using this mode, vba-edit will check that the office document is not already open in another instance of the application. If the document is already open in another instance, the tool will notify the user and exit gracefully.

## PowerQuery M Code

- PowerQuery M code is stored in a different part of the Office (Excel) document than VBA code. PowerQuery M code is stored in the `xl/query` folder of the `.xlsx` file.
- To export/import PowerQuery M code, the tool will need to read/write the relevant parts of the Office document. This can be done using libraries such as `zipfile` to manipulate the `.xlsx` file directly.


## Virtual Filesystem

- The most elegant solution to keep the file system and the VBA editor in sync would be to implement a virtual filesystem that mirrors the structure of the VBA project.
- This would allow users to interact with the VBA project as if it were a regular file system, making it easier to manage and edit VBA code.
- Libraries such as `pyfilesystem2` or `fusepy` could be used to implement this feature.


## Draft Plan: VS Code VBA Virtual Filesystem Extension
possible references:
https://code.visualstudio.com/api/references/vscode-api#FileSystemProvider
https://github.com/microsoft/vscode-extension-samples/tree/main/fsprovider-sample
https://github.com/appulate/vscode-file-watcher

### 1. **Goals**
- Present the VBA project structure (folders, modules, forms) as a virtual file tree in VS Code.
- Allow editing VBA modules directly in VS Code.
- Sync changes between the Office document and the file system automatically.

### 2. **Core Components**
- **FileSystemProvider:** Implements VS Code’s virtual filesystem API.
- **VBA Project Bridge:** Python backend (using the existing `vba-edit` logic) to read/write VBA modules.
- **Extension Host:** TypeScript/JavaScript code that communicates with the Python backend.

### 3. **Development Steps**

#### A. VS Code Extension Skeleton
- Use `yo code` to scaffold a new extension:
  ```sh
  npm install -g yo generator-code
  yo code
  ```
- Choose “New Extension (TypeScript)” and set up your project.

#### B. Implement FileSystemProvider
- In your extension, create a class that implements [`FileSystemProvider`](https://code.visualstudio.com/api/extension-guides/virtual-filesystem).
- This class will handle file/folder listing, reading, writing, and change events.

#### C. Python Backend
- Expose your `vba-edit` logic via a simple REST API, WebSocket, or CLI interface.
- The backend should support:
  - Listing modules/folders
  - Reading module code
  - Writing module code
  - Syncing changes to the Office document

#### D. Communication Layer
- Use Node.js `child_process` to call Python scripts, or use HTTP/WebSocket for more advanced scenarios.
- The extension host will call the backend to fetch/update VBA project data.

#### E. UI Integration
- Register the virtual filesystem as a new workspace or explorer view.
- Optionally, add commands for “Sync”, “Export”, “Import”, etc.

#### F. Testing & Packaging
- Test with sample Office documents.
- Package and publish the extension to the VS Code Marketplace.

---

## Example: FileSystemProvider Skeleton (TypeScript)

````typescript
import * as vscode from 'vscode';

export class VbaFsProvider implements vscode.FileSystemProvider {
    // Implement required methods: stat, readDirectory, readFile, writeFile, etc.
    // Use Python backend to fetch/update VBA modules

    onDidChangeFile: vscode.Event<vscode.FileChangeEvent[]>;

    // ...implementation...
}
````

---

## Next Steps

1. Scaffold the extension (`yo code`).
2. Implement the `FileSystemProvider` skeleton.
3. Connect to your Python backend for VBA project operations.
4. Iterate and test with real Office documents.

---

**Let me know which part you want to start with, or if you want a more detailed code example for any step!**