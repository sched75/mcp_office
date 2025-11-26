"""
Tests d'intégration pour Excel avec application réelle.

ATTENTION : Ces tests interagissent avec Microsoft Excel installé sur le système.
Ils créent et suppriment des fichiers réels dans un dossier temporaire.
"""

import contextlib
import tempfile
from pathlib import Path

import pytest

from src.excel.excel_service import ExcelService


class TestExcelIntegration:
    """Tests d'intégration Excel avec application réelle."""

    @pytest.fixture(scope="class")
    def excel_service(self):
        """Initialise le service Excel."""
        service = ExcelService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire pour les tests."""
        temp_path = Path(tempfile.mkdtemp(prefix="excel_test_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_create_and_save_workbook(self, excel_service, temp_dir):
        """Test : Créer et sauvegarder un classeur."""
        # Créer
        result = excel_service.create_workbook()
        assert result["success"]

        # Écrire des données
        result = excel_service.write_cell("Sheet1", "A1", "Test Excel")
        assert result["success"]

        # Sauvegarder
        output_path = temp_dir / "test_workbook.xlsx"
        result = excel_service.save_workbook(str(output_path))
        assert result["success"]
        assert output_path.exists()

        # Fermer
        excel_service.close_workbook(save=False)

    def test_open_existing_workbook(self, excel_service, temp_dir):
        """Test : Ouvrir un classeur existant."""
        # Créer un classeur d'abord
        excel_service.create_workbook()
        excel_service.write_cell("Sheet1", "A1", "Données de test")
        workbook_path = temp_dir / "existing.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        # Ouvrir
        result = excel_service.open_workbook(str(workbook_path))
        assert result["success"]

        # Vérifier qu'on peut lire les données
        result = excel_service.read_cell("Sheet1", "A1")
        assert result["success"]
        assert result.get("value") == "Données de test"

        excel_service.close_workbook(save=False)

    def test_formulas_and_calculations(self, excel_service, temp_dir):
        """Test : Formules et calculs."""
        excel_service.create_workbook()

        # Écrire des données
        excel_service.write_cell("Sheet1", "A1", "10")
        excel_service.write_cell("Sheet1", "A2", "20")
        excel_service.write_cell("Sheet1", "A3", "30")

        # Utiliser une formule
        result = excel_service.use_function("Sheet1", "A4", "SUM", "A1:A3")
        assert result["success"]

        # Sauvegarder
        workbook_path = temp_dir / "with_formulas.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        assert workbook_path.exists()

    def test_create_chart(self, excel_service, temp_dir):
        """Test : Création de graphique."""
        excel_service.create_workbook()

        # Écrire des données pour le graphique
        data = [["Mois", "Ventes"], ["Janvier", 100], ["Février", 150], ["Mars", 200]]
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                excel_service.write_cell("Sheet1", f"{chr(65 + j)}{i + 1}", value)

        # Créer un graphique
        result = excel_service.create_chart("Sheet1", "Column", "A1:B4")
        assert result["success"]

        # Sauvegarder
        workbook_path = temp_dir / "with_chart.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        assert workbook_path.exists()

    def test_formatting_and_styles(self, excel_service, temp_dir):
        """Test : Formatage et styles."""
        excel_service.create_workbook()

        # Écrire et formater
        excel_service.write_cell("Sheet1", "A1", "Texte formaté")
        result = excel_service.set_cell_color("Sheet1", "A1", 255, 200, 200)
        assert result["success"]

        result = excel_service.set_font_color("Sheet1", "A1", 0, 0, 255)
        assert result["success"]

        # Sauvegarder
        workbook_path = temp_dir / "formatted.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        assert workbook_path.exists()

    def test_export_to_pdf(self, excel_service, temp_dir):
        """Test : Export PDF."""
        excel_service.create_workbook()
        excel_service.write_cell("Sheet1", "A1", "Document pour PDF")

        # Sauvegarder d'abord en XLSX
        xlsx_path = temp_dir / "for_pdf.xlsx"
        excel_service.save_workbook(str(xlsx_path))

        # Exporter en PDF
        pdf_path = temp_dir / "exported.pdf"
        result = excel_service.export_to_pdf(str(pdf_path))
        assert result["success"]

        excel_service.close_workbook(save=False)

        # Vérifier que le PDF existe
        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 0

    def test_multiple_worksheets(self, excel_service, temp_dir):
        """Test : Gérer plusieurs feuilles."""
        excel_service.create_workbook()

        # Ajouter des feuilles
        result = excel_service.add_worksheet("Feuille2")
        assert result["success"]

        result = excel_service.add_worksheet("Feuille3")
        assert result["success"]

        # Écrire dans différentes feuilles
        excel_service.write_cell("Sheet1", "A1", "Feuille 1")
        excel_service.write_cell("Feuille2", "A1", "Feuille 2")
        excel_service.write_cell("Feuille3", "A1", "Feuille 3")

        # Sauvegarder
        workbook_path = temp_dir / "multiple_sheets.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        assert workbook_path.exists()


@pytest.mark.slow
class TestExcelAdvancedIntegration:
    """Tests avancés nécessitant plus de temps."""

    @pytest.fixture(scope="class")
    def excel_service(self):
        """Initialise le service Excel."""
        service = ExcelService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire."""
        temp_path = Path(tempfile.mkdtemp(prefix="excel_advanced_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_complex_workbook_workflow(self, excel_service, temp_dir):
        """Test : Workflow complexe de classeur."""
        excel_service.create_workbook()

        # Ajouter contenu varié
        excel_service.write_cell("Sheet1", "A1", "Rapport Financier Q1 2024")
        excel_service.write_cell("Sheet1", "A3", "Revenus")
        excel_service.write_cell("Sheet1", "B3", "100000")
        excel_service.write_cell("Sheet1", "A4", "Dépenses")
        excel_service.write_cell("Sheet1", "B4", "75000")
        excel_service.write_cell("Sheet1", "A5", "Profit")

        # Utiliser formule pour calculer le profit
        excel_service.use_function("Sheet1", "B5", "SUBTRACT", "B3:B4")

        # Créer un graphique
        excel_service.create_chart("Sheet1", "Column", "A3:B5")

        # Sauvegarder
        workbook_path = temp_dir / "rapport_complet.xlsx"
        result = excel_service.save_workbook(str(workbook_path))
        assert result["success"]
        assert workbook_path.exists()

        excel_service.close_workbook(save=False)

    def test_performance_large_dataset(self, excel_service, temp_dir):
        """Test : Performance sur jeu de données volumineux."""
        import time

        excel_service.create_workbook()

        start_time = time.time()

        # Ajouter 100 lignes de données
        for i in range(100):
            excel_service.write_cell("Sheet1", f"A{i + 1}", f"Ligne {i + 1}")
            excel_service.write_cell("Sheet1", f"B{i + 1}", str(i * 10))

        elapsed = time.time() - start_time
        print(f"\nTemps pour 100 lignes : {elapsed:.2f}s")

        # Sauvegarder
        workbook_path = temp_dir / "large_dataset.xlsx"
        excel_service.save_workbook(str(workbook_path))
        excel_service.close_workbook(save=False)

        assert workbook_path.exists()
        # Le classeur devrait être créé en moins de 60 secondes
        assert elapsed < 60


def run_integration_tests():
    """Exécute tous les tests d'intégration Excel."""
    print("=" * 70)
    print("TESTS D'INTÉGRATION EXCEL")
    print("=" * 70)
    print()
    print("⚠️  Ces tests vont interagir avec Microsoft Excel installé.")
    print("Assurez-vous que Excel est installé et fermé avant de continuer.")
    print()

    response = input("Continuer ? (o/N) : ")
    if response.lower() != "o":
        print("Tests annulés.")
        return

    # Exécuter pytest
    pytest.main(
        [
            __file__,
            "-v",
            "--tb=short",
            "-s",  # Afficher les prints
        ]
    )


if __name__ == "__main__":
    run_integration_tests()
