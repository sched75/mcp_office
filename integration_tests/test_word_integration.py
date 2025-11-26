"""
Tests d'intégration pour Word avec application réelle.

ATTENTION : Ces tests interagissent avec Microsoft Word installé sur le système.
Ils créent et suppriment des fichiers réels dans un dossier temporaire.
"""

import contextlib
import tempfile
from pathlib import Path

import pytest

from src.word.word_service import WordService


class TestWordIntegration:
    """Tests d'intégration Word avec application réelle."""

    @pytest.fixture(scope="class")
    def word_service(self):
        """Initialise le service Word."""
        service = WordService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire pour les tests."""
        temp_path = Path(tempfile.mkdtemp(prefix="word_test_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_create_and_save_document(self, word_service, temp_dir):
        """Test : Créer et sauvegarder un document."""
        # Créer
        result = word_service.create_document()
        assert result["success"]

        # Ajouter du contenu
        result = word_service.add_paragraph("Test d'intégration Word")
        assert result["success"]

        # Sauvegarder
        output_path = temp_dir / "test_document.docx"
        result = word_service.save_document(str(output_path))
        assert result["success"]
        assert output_path.exists()

        # Fermer
        word_service.close_document(save=False)

    def test_open_existing_document(self, word_service, temp_dir):
        """Test : Ouvrir un document existant."""
        # Créer un document d'abord
        word_service.create_document()
        word_service.add_paragraph("Document de test")
        doc_path = temp_dir / "existing.docx"
        word_service.save_document(str(doc_path))
        word_service.close_document(save=False)

        # Ouvrir
        result = word_service.open_document(str(doc_path))
        assert result["success"]

        # Vérifier qu'on peut ajouter du contenu
        result = word_service.add_paragraph("Nouveau paragraphe")
        assert result["success"]

        word_service.close_document(save=False)

    def test_text_formatting(self, word_service, temp_dir):
        """Test : Formatage de texte."""
        word_service.create_document()

        # Ajouter et formater
        result = word_service.add_paragraph("Texte en gras")
        assert result["success"]

        result = word_service.apply_text_formatting(bold=True, font_size=14, font_name="Arial")
        assert result["success"]

        # Sauvegarder
        doc_path = temp_dir / "formatted.docx"
        word_service.save_document(str(doc_path))
        word_service.close_document(save=False)

        assert doc_path.exists()

    def test_insert_table(self, word_service, temp_dir):
        """Test : Insertion de tableau."""
        word_service.create_document()

        # Insérer tableau
        result = word_service.insert_table(rows=3, cols=3)
        assert result["success"]

        # Remplir des cellules
        result = word_service.set_table_cell_text(table_index=1, row=1, col=1, text="A1")
        assert result["success"]

        # Sauvegarder
        doc_path = temp_dir / "with_table.docx"
        word_service.save_document(str(doc_path))
        word_service.close_document(save=False)

        assert doc_path.exists()

    def test_find_and_replace(self, word_service, temp_dir):
        """Test : Rechercher et remplacer."""
        word_service.create_document()

        # Ajouter du texte
        word_service.add_paragraph("Hello World. Hello Universe.")

        # Remplacer
        result = word_service.find_and_replace("Hello", "Hi")
        assert result["success"]
        assert result.get("replacements", 0) == 2

        word_service.close_document(save=False)

    def test_export_to_pdf(self, word_service, temp_dir):
        """Test : Export PDF."""
        word_service.create_document()
        word_service.add_paragraph("Document pour PDF")

        # Sauvegarder d'abord en DOCX
        docx_path = temp_dir / "for_pdf.docx"
        word_service.save_document(str(docx_path))

        # Exporter en PDF
        pdf_path = temp_dir / "exported.pdf"
        result = word_service.print_to_pdf(str(pdf_path))
        assert result["success"]

        word_service.close_document(save=False)

        # Vérifier que le PDF existe
        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 0

    def test_multiple_documents(self, word_service, temp_dir):
        """Test : Gérer plusieurs documents."""
        # Créer premier document
        word_service.create_document()
        word_service.add_paragraph("Document 1")
        doc1_path = temp_dir / "doc1.docx"
        word_service.save_document(str(doc1_path))
        word_service.close_document()

        # Créer second document
        word_service.create_document()
        word_service.add_paragraph("Document 2")
        doc2_path = temp_dir / "doc2.docx"
        word_service.save_document(str(doc2_path))
        word_service.close_document()

        assert doc1_path.exists()
        assert doc2_path.exists()


@pytest.mark.slow
class TestWordAdvancedIntegration:
    """Tests avancés nécessitant plus de temps."""

    @pytest.fixture(scope="class")
    def word_service(self):
        """Initialise le service Word."""
        service = WordService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire."""
        temp_path = Path(tempfile.mkdtemp(prefix="word_advanced_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_complex_document_workflow(self, word_service, temp_dir):
        """Test : Workflow complexe de document."""
        word_service.create_document()

        # Ajouter contenu varié
        word_service.add_paragraph("Rapport Annuel 2024", style="Heading 1")
        word_service.add_paragraph("Résumé Exécutif", style="Heading 2")
        word_service.add_paragraph("Ce rapport présente les résultats de l'année 2024...")

        # Tableau de données
        word_service.insert_table(rows=5, cols=3)
        word_service.set_table_cell_text(1, 1, 1, "Trimestre")
        word_service.set_table_cell_text(1, 1, 2, "Ventes")
        word_service.set_table_cell_text(1, 1, 3, "Profit")

        # Ajouter header et footer
        word_service.add_header("Confidentiel - 2024")
        word_service.add_footer("Page")

        # Sauvegarder
        doc_path = temp_dir / "rapport_complet.docx"
        result = word_service.save_document(str(doc_path))
        assert result["success"]
        assert doc_path.exists()

        word_service.close_document(save=False)

    def test_performance_large_document(self, word_service, temp_dir):
        """Test : Performance sur document volumineux."""
        import time

        word_service.create_document()

        start_time = time.time()

        # Ajouter 100 paragraphes
        for i in range(100):
            word_service.add_paragraph(f"Paragraphe {i + 1} " * 10)

        elapsed = time.time() - start_time
        print(f"\nTemps pour 100 paragraphes : {elapsed:.2f}s")

        # Sauvegarder
        doc_path = temp_dir / "large_document.docx"
        word_service.save_document(str(doc_path))
        word_service.close_document(save=False)

        assert doc_path.exists()
        # Le document devrait être créé en moins de 60 secondes
        assert elapsed < 60


def run_integration_tests():
    """Exécute tous les tests d'intégration Word."""
    print("=" * 70)
    print("TESTS D'INTÉGRATION WORD")
    print("=" * 70)
    print()
    print("⚠️  Ces tests vont interagir avec Microsoft Word installé.")
    print("Assurez-vous que Word est installé et fermé avant de continuer.")
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
