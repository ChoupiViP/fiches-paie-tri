import pytest

import trier_fiches_paie as c


def test_en_nombre():
    assert c.en_nombre("1 234,56") == 1234.56
    assert c.en_nombre("1.234,56") == 1234.56
    assert c.en_nombre("1234.56") == 1234.56
    assert c.en_nombre("-12,50") == -12.5


def test_montants_ligne():
    assert c.montants_ligne("Net imposable 1 432,18") == [1432.18]
    assert c.montants_ligne("Net imposable 1 432,18 15 000,00") == [1432.18, 15000.0]
    assert c.montants_ligne("Période du 01/03/2026 au 31/03/2026") == []


@pytest.mark.parametrize("texte, attendu", [
    ("Période du 01/03/2026 au 31/03/2026", (2026, 3)),
    ("Période : 07/2024", (2024, 7)),
    ("Bulletin de salaire - mois de décembre 2025", (2025, 12)),
    ("Bulletin de paie mars 2023", (2023, 3)),
    ("rien d'utile ici", None),
])
def test_periode(texte, attendu):
    assert c.detecter_periode(texte) == attendu


def test_employeur():
    texte = "Employeur : SARL Dupont Informatique\nautre ligne"
    assert c.detecter_employeur(texte, texte.splitlines()) == "SARL Dupont Informatique"
    texte = "ADECCO France\nBulletin de paie"
    assert c.detecter_employeur(texte, texte.splitlines()) == "Adecco"


def _pdf(chemin, lignes):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    cv = canvas.Canvas(str(chemin), pagesize=A4)
    y = 800
    for ligne in lignes:
        cv.drawString(40, y, ligne)
        y -= 18
    cv.save()


def test_bout_en_bout(tmp_path):
    pytest.importorskip("reportlab")
    src = tmp_path / "pdf"
    src.mkdir()
    # trois mises en page différentes
    _pdf(src / "a1.pdf", ["ADECCO France", "BULLETIN DE PAIE", "Période du 01/03/2026 au 31/03/2026",
                          "Salaire brut 1 850,00", "Net imposable 1 432,18"])
    _pdf(src / "a2.pdf", ["ADECCO France", "BULLETIN DE PAIE", "Période du 01/05/2026 au 31/05/2026",
                          "Salaire brut 1 900,00", "Net imposable 1 470,00",
                          "Net imposable cumulé 2 902,18"])
    _pdf(src / "b1.pdf", ["Employeur : SARL Dupont Informatique", "Bulletin de salaire - mois de octobre 2025",
                          "Total brut 2 100,50", "Net fiscal 1 650,25"])
    _pdf(src / "c1.pdf", ["Société Martin Frères", "Période : 07/2024", "Net imposable", "1 210,00"])
    # doublon exact
    (src / "a1_copie.pdf").write_bytes((src / "a1.pdf").read_bytes())

    msgs = []
    res = c.traiter(src, tmp_path / "out", log=msgs.append)

    assert res["nb_pdf"] == 5
    assert res["doublons"] == 1
    assert res["lues"] == 4
    assert res["totaux"][2026] == pytest.approx(2902.18)
    assert res["totaux"][2025] == pytest.approx(1650.25)
    assert res["totaux"][2024] == pytest.approx(1210.0)
    assert (tmp_path / "out" / "2026" / "2026-03_Adecco.pdf").exists()
    assert res["xlsx"].exists()


def test_dossier_sans_pdf(tmp_path):
    with pytest.raises(ValueError):
        c.traiter(tmp_path, tmp_path / "out", log=lambda m: None)


def test_dossier_introuvable(tmp_path):
    with pytest.raises(ValueError):
        c.traiter(tmp_path / "nexiste_pas", tmp_path / "out", log=lambda m: None)
