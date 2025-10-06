`vba-edit <https://github.com/markuskiller/vba-edit>`__
=======================================================

**Edit VBA code in VS Code, Sublime, or any editor you love.** Real-time sync with MS Office apps (support for **Excel**, **Word**, **PowerPoint** & **Access**). Git-friendly. No more VBA editor pain.

|CI| |PyPI - Version| |PyPI - Python Version| |PyPI - Downloads| |License|

.. _30-second-demo:

30-Second Demo
--------------

.. code:: bash

   # Install
   pip install -U vba-edit

   # Start editing (uses active Excel/Word document)
   excel-vba edit    # or word-vba edit

   # That's it! Edit the .bas/.cls files in your editor. Save = Sync.

How It Works
------------

.. raw:: html

   <pre>
                            ←── vba-edit ──→

   Excel / Word                 COMMANDS            Your favourite
   PowerPoint / Access              ↓                   Editor

   ┌──────────────────┐                          ┌──────────────────┐
   │                  │                          │                  │
   │   VBA Project    │    ←──    EDIT* (once →) │  (e.g. VS CODE)  │ 
   │                  │                          │                  │     latest
   │  (Office VBA-    │           EXPORT ──→     │   .bas           │   ← AI coding-  
   │    Editor)       │                          │   .cls           │     assistants
   │                  │       ←── IMPORT         │   .frm           │   
   │                  │                          │  (.frx binary)   │ 
   │                  │                          │                  │ 
   └──────────────────┘                          └──────────────────┘
                                                          ↓
                                                 ┌──────────────────┐
                                                 │                  │
    * watches & syncs                            │    (e.g. Git)    │
      back to Office                             │  version control │
      VBA-Editor live                            │                  │
      on save [CTRL+S]                           │                  │
                                                 └──────────────────┘
   </pre>

Why vba-edit?
-------------

- **Use YOUR editor** - VS Code, PyCharm, Wing IDE, Vim, etc. whatever you love
- **AI-ready** - Use Copilot, ChatGPT, or any coding assistant
- **Team-friendly** - Share code via Git, no COM add-ins needed
- **Real version control** - Diff, merge, and track changes properly
- **Well-organized** - Keep your VBA structured, clean, and consistent

Setup (One-Time)
----------------

**Windows Only** \| **MS Office**

Enable VBA access in Office:

``File → Options → Trust Center → Trust Center Settings → Macro Settings``

✅ **Trust access to the VBA project object model**

   💡 Can't find it? Run ``excel-vba check`` to verify settings

Common Workflows
----------------

Start Fresh
~~~~~~~~~~~

.. code:: bash

   excel-vba edit                    # Start with active workbook

Quick Export with Folder View
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: bash

   excel-vba export --open-folder    # Export and open in File Explorer

Team Project with Git
~~~~~~~~~~~~~~~~~~~~~

.. code:: bash

   excel-vba export --vba-directory ./src/vba
   git add . && git commit -m "Updated reports module"

Support for RubberduckVBA Style (big thank you to @onderhold!)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: bash

   excel-vba edit --rubberduck-folders --in-file-headers

Quick Reference
---------------

+----------------------------------------+-----------------------------------------------------------------+
| Command                                | What it does                                                    |
+========================================+=================================================================+
| ``excel-vba edit``                     | Start live editing                                              |
+----------------------------------------+-----------------------------------------------------------------+
| ``excel-vba export``                   | One-time export                                                 |
+----------------------------------------+-----------------------------------------------------------------+
| ``excel-vba export --open-folder``     | Export and open folder in explorer                              |
+----------------------------------------+-----------------------------------------------------------------+
| ``excel-vba export --force-overwrite`` | Export without confirmation prompts                             |
+----------------------------------------+-----------------------------------------------------------------+
| ``excel-vba check``                    | Verify status of *Trust access* to the VBA project object model |
+----------------------------------------+-----------------------------------------------------------------+
| ``--vba-directory ./src``              | Custom folder                                                   |
+----------------------------------------+-----------------------------------------------------------------+
| ``--rubberduck-folders``               | Organize by @Folder                                             |
+----------------------------------------+-----------------------------------------------------------------+
| ``--in-file-headers``                  | Embed headers in code files                                     |
+----------------------------------------+-----------------------------------------------------------------+
| ``--conf myconfig.toml``               | Use config file                                                 |
+----------------------------------------+-----------------------------------------------------------------+
| ``--force-overwrite``                  | Skip safety prompts (automation)                                |
+----------------------------------------+-----------------------------------------------------------------+

Troubleshooting
---------------

==================== =======================================
Issue                Solution
==================== =======================================
"Trust access" error Run ``excel-vba check`` for diagnostics
Changes not syncing  Save the file in your editor
Forms not working    Add ``--in-file-headers`` flag
==================== =======================================

Safety Features
---------------

**🛡️ Data Loss Prevention** (v0.4.0+)

vba-edit now protects your work with smart safety checks:

- **Overwrite Protection**: Warns before overwriting existing VBA files
- **Header Mode Detection**: Alerts when switching between header storage modes
- **Orphaned File Cleanup**: Automatically removes stale ``.header`` files on mode change
- **UserForm Validation**: Prevents exports without proper header handling

**Bypass for Automation**: Use ``--force-overwrite`` flag to skip prompts in CI/CD pipelines:

.. code:: bash

   excel-vba export --vba-directory ./src --force-overwrite

..

   ⚠️ **CAUTION**: ``--force-overwrite`` suppresses all safety prompts. Use with caution!

Features
--------

**🚀 Core**

- Live sync between Office and your editor
- Full Git/version control support
- All Office apps: Excel, Word, Access & **NEW v0.4.0+** PowerPoint

**📁 Organization**

- **NEW v0.4.0+** RubberduckVBA folder structure support
- **NEW v0.4.0+** Smart file organization with ``@Folder`` annotations
- **NEW v0.4.0+** TOML config files for team standards

**🔧 Advanced**

- Unicode & encoding support
- **IMPROVED v0.4.0+** UserForms with layout preservation
- Class modules with custom attributes

Command Line Tools
------------------

App-specific tools
~~~~~~~~~~~~~~~~~~

- ``word-vba``
- ``excel-vba``
- ``access-vba``
- ``powerpoint-vba``

Commands
~~~~~~~~

- ``edit``: Live sync between editor and Office
- ``export``: Export VBA modules to files
- ``import``: Import VBA modules from files
- ``check {all}``: Check if 'Trust Access to the VBA project object model' is enabled

Options
~~~~~~~

.. code:: text

   --file, -f                   Path to Office document
   --conf, --config, -c         Supply config file (TOML format)
   --vba-directory              Directory for VBA files
   --rubberduck-folders         Use RubberduckVBA folder annotations
   --save-headers               Save module headers separately
   --in-file-headers            Include VBA headers directly in code files
   --encoding, -e               Specify character encoding
   --detect-encoding, -d        Auto-detect encoding
   --verbose, -v                Enable detailed logging
   --logfile, -l                Enable file logging
   --open-folder                Open export directory in file explorer after export
   --save-metadata, -m          Save metadata file with encoding information
   --force-overwrite            Skip all confirmation prompts (for automation)
   --version, -V                Show program's version number and exit

.. _example-of---in-file-headers---rubberduck-folders-v040:

Example of ``--in-file-headers --rubberduck-folders`` (v0.4.0+)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: vba

   VERSION 1.0 CLASS
   BEGIN
     MultiUse = -1  'True
   END
   Attribute VB_Name = "MyClass"
   Attribute VB_GlobalNameSpace = False
   Attribute VB_Creatable = False
   Attribute VB_PredeclaredId = False
   Attribute VB_Exposed = False

   '@Folder("Business.Domain")
   Public Sub DoSomething()
       ' Your code here
   End Sub

Configuration Files
-------------------

**NEW v0.4.0+** Use TOML configuration files to standardize team workflows and avoid repetitive command-line arguments.

Basic Configuration
~~~~~~~~~~~~~~~~~~~

Create a ``vba-config.toml`` file in your project:

.. code:: toml

   [general]
   file = "MyWorkbook.xlsm"
   vba_directory = "src/vba"
   verbose = true
   rubberduck_folders = true
   in_file_headers = true

Then use it:

.. code:: bash

   excel-vba export --conf vba-config.toml

Available Configuration Keys
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**[general] section:**

- ``file`` - Path to Office document
- ``vba_directory`` - Directory for VBA files
- ``encoding`` - Character encoding (e.g., "utf-8", "cp1252")
- ``verbose`` - Enable verbose logging (true/false)
- ``logfile`` - Path to log file
- ``rubberduck_folders`` - Use RubberduckVBA @Folder annotations (true/false)
- ``save_headers`` - Save headers to separate .header files (true/false)
- ``in_file_headers`` - Embed headers in code files (true/false)
- ``open_folder`` - Open export directory after export (true/false)

**Other sections (reserved for future use):**

- ``[office]`` - Office-wide settings
- ``[excel]`` - Excel-specific settings
- ``[word]`` - Word-specific settings
- ``[access]`` - Access-specific settings
- ``[powerpoint]`` - PowerPoint-specific settings

Configuration Placeholders
~~~~~~~~~~~~~~~~~~~~~~~~~~

Configuration values support dynamic placeholders for flexible path management:

**Available placeholders:**

- ``{config.path}`` - Directory containing the config file
- ``{general.file.name}`` - Document filename without extension
- ``{general.file.fullname}`` - Document filename with extension
- ``{general.file.path}`` - Directory containing the document
- ``{vbaproject}`` - VBA project name (resolved at runtime)

**Example with placeholders:**

.. code:: toml

   [general]
   file = "C:/Projects/MyApp/MyWorkbook.xlsm"
   vba_directory = "{general.file.path}/{general.file.name}-vba"
   # This resolves to: C:/Projects/MyApp/MyWorkbook-vba

**Relative paths example:**

.. code:: toml

   [general]
   file = "../documents/report.xlsm"
   vba_directory = "{config.path}/vba-modules"
   # vba_directory is relative to config file location

Command-Line Override
~~~~~~~~~~~~~~~~~~~~~

Command-line arguments always override config file settings:

.. code:: bash

   # Config says vba_directory = "src/vba"
   # This overrides it to "build/vba"
   excel-vba export --conf vba-config.toml --vba-directory build/vba

..

   ⚠️ **CAUTION**: **1.** Always **backup your Office files** before using ``vba-edit`` **2.** Use **version control (git)** to track your VBA code **3.** Run ``export`` after changing **form layouts** or module properties

Known Limitations
~~~~~~~~~~~~~~~~~

- UserForms require ``--save-headers`` option (``edit`` process is aborted if this is not the case)
- If ``*.header`` files are modified on their own, the corresponding ``*.cls``, ``*.bas`` or ``*.frm`` file needs to be saved in order to sync the complete module back into the VBA project model

Links
-----

- `Homepage <https://langui.ch/current-projects/vba-edit/>`__
- `Documentation <https://github.com/markuskiller/vba-edit/blob/main/README.md>`__
- `Source Code <https://github.com/markuskiller/vba-edit>`__
- `Changelog <https://github.com/markuskiller/vba-edit/blob/main/CHANGELOG.md>`__
- `Changelog of latest dev version <https://github.com/markuskiller/vba-edit/blob/dev/CHANGELOG.md>`__
- `Video Tutorial <https://www.youtube.com/watch?v=xoO-Fx0fTpM>`__ (xlwings walkthrough, with similar functionality)

License
-------

BSD 3-Clause License

Credits
-------

**vba-edit** builds on an excellent idea first implemented for Excel in `xlwings <https://www.xlwings.org/>`__ (BSD-3).

Special thanks to **@onderhold** for improved header handling, RubberduckVBA folder and config file support in v0.4.0.

.. |CI| image:: https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg
   :target: https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml
.. |PyPI - Version| image:: https://img.shields.io/pypi/v/vba-edit.svg
   :target: https://pypi.org/project/vba-edit
.. |PyPI - Python Version| image:: https://img.shields.io/pypi/pyversions/vba-edit.svg
   :target: https://pypi.org/project/vba-edit
.. |PyPI - Downloads| image:: https://img.shields.io/pypi/dm/vba-edit
   :target: https://pypi.org/project/vba-edit
.. |License| image:: https://img.shields.io/badge/License-BSD_3--Clause-blue.svg
   :target: https://opensource.org/licenses/BSD-3-Clause
