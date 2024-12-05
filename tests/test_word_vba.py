import subprocess

def test_word_vba_help():
    result = subprocess.run(['word-vba', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: word-vba" in result.stdout
    assert "Commands:" in result.stdout
    assert "edit" in result.stdout
    assert "import" in result.stdout
    assert "export" in result.stdout

def test_word_vba_edit():
    result = subprocess.run(['word-vba', 'edit', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: word-vba edit" in result.stdout
    assert "--verbose" in result.stdout


def test_word_vba_import():
    result = subprocess.run(['word-vba', 'import', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: word-vba import" in result.stdout
    assert "--verbose" in result.stdout

def test_word_vba_export():
    result = subprocess.run(['word-vba', 'export', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: word-vba export" in result.stdout
    assert "--verbose" in result.stdout