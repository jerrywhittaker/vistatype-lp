"""inject_customui.py -- putting the ribbon inside the built .dotm.

This is the last step of `make build`, after Word has produced the .dotm. Re-running it must
replace the ribbon part and its relationship, never duplicate them: a second relationship
pointing at the same part is the kind of thing Word tolerates until it does not.
"""
import subprocess
import sys
import zipfile

import inject_customui as mod

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\
</Relationships>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
<Default Extension="xml" ContentType="application/xml"/>\
</Types>"""


def count_customui_rels(rels_xml):
    return rels_xml.count('Target="customUI/customUI14.xml"')


def test_the_relationship_is_added():
    out = mod.add_relationship(RELS)
    assert count_customui_rels(out) == 1
    assert mod.REL_ID in out
    assert out.endswith("</Relationships>")


def test_the_existing_document_relationship_survives():
    out = mod.add_relationship(RELS)
    assert 'Target="word/document.xml"' in out


def test_adding_it_twice_leaves_one(): 
    out = mod.add_relationship(mod.add_relationship(RELS))
    assert count_customui_rels(out) == 1


def test_a_relationship_written_by_word_with_a_different_id_is_replaced():
    """Word may have put its own there; ours must not end up beside it."""
    existing = RELS.replace(
        "</Relationships>",
        '<Relationship Id="rId99" Type="x" Target="customUI/customUI14.xml"/>'
        "</Relationships>")
    out = mod.add_relationship(existing)
    assert count_customui_rels(out) == 1
    assert "rId99" not in out


def dotm(tmp_path):
    path = tmp_path / "LPandBRL.dotm"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", "<document/>")
        z.writestr("word/vbaProject.bin", b"\x00binary\x00")
    return path


def ribbon(tmp_path, text='<customUI><ribbon/></customUI>'):
    path = tmp_path / "customUI14.xml"
    path.write_text(text, encoding="utf-8")
    return path


def inject(tmp_path, path, rib):
    return subprocess.run(
        [sys.executable, str(mod.__file__), str(path), str(rib)],
        cwd=str(tmp_path), capture_output=True, text=True)


def test_the_ribbon_lands_in_the_package(tmp_path):
    path, rib = dotm(tmp_path), ribbon(tmp_path)
    r = inject(tmp_path, path, rib)
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(path) as z:
        assert z.read(mod.PART).decode("utf-8") == rib.read_text()
        assert count_customui_rels(z.read("_rels/.rels").decode("utf-8")) == 1


def test_nothing_else_in_the_package_is_disturbed(tmp_path):
    path, rib = dotm(tmp_path), ribbon(tmp_path)
    inject(tmp_path, path, rib)
    with zipfile.ZipFile(path) as z:
        assert z.read("word/vbaProject.bin") == b"\x00binary\x00"
        assert z.read("word/document.xml").decode("utf-8") == "<document/>"


def test_content_types_stays_first_in_the_package(tmp_path):
    """Conventional, and some readers assume it."""
    path, rib = dotm(tmp_path), ribbon(tmp_path)
    inject(tmp_path, path, rib)
    with zipfile.ZipFile(path) as z:
        assert z.namelist()[0] == "[Content_Types].xml"


def test_injecting_twice_replaces_rather_than_duplicates(tmp_path):
    path, rib = dotm(tmp_path), ribbon(tmp_path)
    inject(tmp_path, path, rib)
    ribbon(tmp_path, '<customUI><ribbon><tabs/></ribbon></customUI>')
    inject(tmp_path, path, rib)
    with zipfile.ZipFile(path) as z:
        assert z.namelist().count(mod.PART) == 1
        assert count_customui_rels(z.read("_rels/.rels").decode("utf-8")) == 1
        assert "<tabs/>" in z.read(mod.PART).decode("utf-8")


def test_a_package_with_no_rels_part_is_refused(tmp_path):
    path = tmp_path / "broken.dotm"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", "<document/>")
    r = inject(tmp_path, path, ribbon(tmp_path))
    assert r.returncode != 0
    assert "_rels/.rels not found" in r.stderr
