"""
Tests d'intégration pour PowerPoint avec application réelle.

ATTENTION : Ces tests interagissent avec Microsoft PowerPoint installé sur le système.
Ils créent et suppriment des fichiers réels dans un dossier temporaire.
"""

import contextlib
import tempfile
from pathlib import Path

import pytest

from src.powerpoint.powerpoint_service import PowerPointService


class TestPowerPointIntegration:
    """Tests d'intégration PowerPoint avec application réelle."""

    @pytest.fixture(scope="class")
    def powerpoint_service(self):
        """Initialise le service PowerPoint."""
        service = PowerPointService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire pour les tests."""
        temp_path = Path(tempfile.mkdtemp(prefix="powerpoint_test_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_create_and_save_presentation(self, powerpoint_service, temp_dir):
        """Test : Créer et sauvegarder une présentation."""
        # Créer
        result = powerpoint_service.create_presentation()
        assert result["success"]

        # Ajouter une diapositive
        result = powerpoint_service.add_slide()
        assert result["success"]

        # Sauvegarder
        output_path = temp_dir / "test_presentation.pptx"
        result = powerpoint_service.save_presentation(str(output_path))
        assert result["success"]
        assert output_path.exists()

        # Fermer
        powerpoint_service.close_presentation(save=False)

    def test_open_existing_presentation(self, powerpoint_service, temp_dir):
        """Test : Ouvrir une présentation existante."""
        # Créer une présentation d'abord
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()
        presentation_path = temp_dir / "existing.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        # Ouvrir
        result = powerpoint_service.open_presentation(str(presentation_path))
        assert result["success"]

        # Vérifier qu'on peut ajouter du contenu
        result = powerpoint_service.add_slide()
        assert result["success"]

        powerpoint_service.close_presentation(save=False)

    def test_slide_management(self, powerpoint_service, temp_dir):
        """Test : Gestion des diapositives."""
        powerpoint_service.create_presentation()

        # Ajouter plusieurs diapositives
        result = powerpoint_service.add_slide()
        assert result["success"]

        result = powerpoint_service.add_slide()
        assert result["success"]

        # Modifier le titre d'une diapositive
        result = powerpoint_service.modify_title(1, "Nouveau Titre")
        assert result["success"]

        # Sauvegarder
        presentation_path = temp_dir / "with_slides.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()

    def test_text_content(self, powerpoint_service, temp_dir):
        """Test : Contenu textuel."""
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()

        # Ajouter une zone de texte
        result = powerpoint_service.add_textbox(1, "Texte de test", 100, 100, 200, 50)
        assert result["success"]

        # Ajouter des puces
        result = powerpoint_service.add_bullets(1, ["Point 1", "Point 2", "Point 3"])
        assert result["success"]

        # Sauvegarder
        presentation_path = temp_dir / "with_text.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()

    def test_insert_image(self, powerpoint_service, temp_dir):
        """Test : Insertion d'image."""
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()

        # Créer une image de test simple (carré rouge)
        from PIL import Image

        test_image_path = temp_dir / "test_image.png"
        img = Image.new("RGB", (100, 100), color="red")
        img.save(test_image_path)

        # Insérer l'image
        result = powerpoint_service.insert_image(1, str(test_image_path), 100, 100)
        assert result["success"]

        # Sauvegarder
        presentation_path = temp_dir / "with_image.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()

    def test_shapes_and_objects(self, powerpoint_service, temp_dir):
        """Test : Formes et objets."""
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()

        # Insérer une forme
        result = powerpoint_service.insert_shape(1, "Rectangle", 100, 100, 200, 100)
        assert result["success"]

        # Modifier la couleur de remplissage
        result = powerpoint_service.modify_fill_color(1, 1, 255, 200, 200)
        assert result["success"]

        # Sauvegarder
        presentation_path = temp_dir / "with_shapes.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()

    def test_export_to_pdf(self, powerpoint_service, temp_dir):
        """Test : Export PDF."""
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()
        powerpoint_service.add_textbox(1, "Présentation pour PDF", 100, 100, 200, 50)

        # Sauvegarder d'abord en PPTX
        pptx_path = temp_dir / "for_pdf.pptx"
        powerpoint_service.save_presentation(str(pptx_path))

        # Exporter en PDF
        pdf_path = temp_dir / "exported.pdf"
        result = powerpoint_service.export_to_pdf(str(pdf_path))
        assert result["success"]

        powerpoint_service.close_presentation(save=False)

        # Vérifier que le PDF existe
        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 0

    def test_transitions_and_animations(self, powerpoint_service, temp_dir):
        """Test : Transitions et animations."""
        powerpoint_service.create_presentation()
        powerpoint_service.add_slide()

        # Appliquer une transition
        result = powerpoint_service.apply_transition(1, "Fade")
        assert result["success"]

        # Ajouter une animation d'entrée
        powerpoint_service.add_textbox(1, "Texte animé", 100, 100, 200, 50)
        result = powerpoint_service.add_entrance_animation(1, 1, "Fly In")
        assert result["success"]

        # Sauvegarder
        presentation_path = temp_dir / "with_animations.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()


@pytest.mark.slow
class TestPowerPointAdvancedIntegration:
    """Tests avancés nécessitant plus de temps."""

    @pytest.fixture(scope="class")
    def powerpoint_service(self):
        """Initialise le service PowerPoint."""
        service = PowerPointService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire."""
        temp_path = Path(tempfile.mkdtemp(prefix="powerpoint_advanced_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_complex_presentation_workflow(self, powerpoint_service, temp_dir):
        """Test : Workflow complexe de présentation."""
        powerpoint_service.create_presentation()

        # Ajouter plusieurs diapositives avec contenu varié
        powerpoint_service.add_slide()
        powerpoint_service.modify_title(1, "Page de Titre")
        powerpoint_service.add_textbox(1, "Présentation MCP Office", 100, 150, 400, 50)

        powerpoint_service.add_slide()
        powerpoint_service.modify_title(2, "Introduction")
        powerpoint_service.add_bullets(
            2, ["Présentation du projet", "Objectifs et fonctionnalités", "Architecture technique"]
        )

        powerpoint_service.add_slide()
        powerpoint_service.modify_title(3, "Démonstration")
        powerpoint_service.add_textbox(3, "Démonstration des fonctionnalités", 100, 150, 400, 50)

        # Appliquer des transitions
        powerpoint_service.apply_transition_to_all("Fade", 1.0)

        # Sauvegarder
        presentation_path = temp_dir / "presentation_complete.pptx"
        result = powerpoint_service.save_presentation(str(presentation_path))
        assert result["success"]
        assert presentation_path.exists()

        powerpoint_service.close_presentation(save=False)

    def test_performance_multiple_slides(self, powerpoint_service, temp_dir):
        """Test : Performance sur présentation volumineuse."""
        import time

        powerpoint_service.create_presentation()

        start_time = time.time()

        # Ajouter 20 diapositives
        for i in range(20):
            powerpoint_service.add_slide()
            powerpoint_service.modify_title(i + 1, f"Diapositive {i + 1}")
            powerpoint_service.add_textbox(i + 1, f"Contenu diapositive {i + 1}", 100, 150, 400, 50)

        elapsed = time.time() - start_time
        print(f"\nTemps pour 20 diapositives : {elapsed:.2f}s")

        # Sauvegarder
        presentation_path = temp_dir / "large_presentation.pptx"
        powerpoint_service.save_presentation(str(presentation_path))
        powerpoint_service.close_presentation(save=False)

        assert presentation_path.exists()
        # La présentation devrait être créée en moins de 60 secondes
        assert elapsed < 60


def run_integration_tests():
    """Exécute tous les tests d'intégration PowerPoint."""
    print("=" * 70)
    print("TESTS D'INTÉGRATION POWERPOINT")
    print("=" * 70)
    print()
    print("⚠️  Ces tests vont interagir avec Microsoft PowerPoint installé.")
    print("Assurez-vous que PowerPoint est installé et fermé avant de continuer.")
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
