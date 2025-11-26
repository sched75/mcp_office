"""
Tests d'intégration pour Outlook avec application réelle.

ATTENTION : Ces tests interagissent avec Microsoft Outlook installé sur le système.
Ils créent et suppriment des éléments réels dans Outlook (emails, contacts, rendez-vous).
"""

import contextlib
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import pytest

from src.outlook.outlook_service import OutlookService


class TestOutlookIntegration:
    """Tests d'intégration Outlook avec application réelle."""

    @pytest.fixture(scope="class")
    def outlook_service(self):
        """Initialise le service Outlook."""
        service = OutlookService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire pour les tests."""
        temp_path = Path(tempfile.mkdtemp(prefix="outlook_test_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_create_and_send_email(self, outlook_service):
        """Test : Créer et envoyer un email."""
        # Créer un email
        result = outlook_service.create_email(
            to="test@example.com",
            subject="Test d'intégration Outlook",
            body="Ceci est un email de test depuis les tests d'intégration.",
        )
        assert result["success"]

        # Ajouter une pièce jointe (fichier temporaire)
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("Contenu de test pour pièce jointe")
            attachment_path = f.name

        try:
            result = outlook_service.add_attachment(attachment_path)
            assert result["success"]

            # Envoyer l'email (ou le sauvegarder en brouillon selon la configuration)
            result = outlook_service.send_email()
            assert result["success"]

        finally:
            # Nettoyer le fichier temporaire
            Path(attachment_path).unlink(missing_ok=True)

    def test_save_email_as_draft(self, outlook_service):
        """Test : Sauvegarder un email en brouillon."""
        result = outlook_service.create_email(
            to="draft@example.com",
            subject="Brouillon de test",
            body="Ceci est un brouillon de test.",
        )
        assert result["success"]

        result = outlook_service.save_as_draft()
        assert result["success"]

    def test_create_contact(self, outlook_service):
        """Test : Créer un contact."""
        result = outlook_service.create_contact(
            first_name="Test",
            last_name="Integration",
            email="test.integration@example.com",
            company="Test Company",
        )
        assert result["success"]

    def test_create_appointment(self, outlook_service):
        """Test : Créer un rendez-vous."""
        start_time = datetime.now() + timedelta(hours=1)
        end_time = start_time + timedelta(hours=2)

        result = outlook_service.create_appointment(
            subject="Rendez-vous de test",
            start_time=start_time.isoformat(),
            end_time=end_time.isoformat(),
            location="Salle de réunion virtuelle",
            body="Ceci est un rendez-vous de test créé depuis les tests d'intégration.",
        )
        assert result["success"]

    def test_folder_operations(self, outlook_service):
        """Test : Opérations sur les dossiers."""
        # Lister les dossiers
        result = outlook_service.list_folders()
        assert result["success"]
        assert "folders" in result

        # Créer un dossier de test
        result = outlook_service.create_folder("Test_Integration_Folder")
        assert result["success"]

        # Vérifier que le dossier existe
        result = outlook_service.list_folders()
        assert result["success"]
        folder_names = [folder["name"] for folder in result["folders"]]
        assert "Test_Integration_Folder" in folder_names

        # Supprimer le dossier de test
        result = outlook_service.delete_folder("Test_Integration_Folder")
        assert result["success"]

    def test_email_search(self, outlook_service):
        """Test : Recherche d'emails."""
        # Rechercher des emails récents
        result = outlook_service.search_emails(subject="test", limit=5)
        assert result["success"]
        # Peut retourner une liste vide si aucun email ne correspond

    def test_calendar_operations(self, outlook_service):
        """Test : Opérations sur le calendrier."""
        # Obtenir les rendez-vous du jour
        today = datetime.now().date()
        result = outlook_service.get_appointments_for_date(today.isoformat())
        assert result["success"]

        # Obtenir les rendez-vous de la semaine
        result = outlook_service.get_appointments_for_week()
        assert result["success"]

    def test_attachment_operations(self, outlook_service, temp_dir):
        """Test : Opérations sur les pièces jointes."""
        # Créer un email avec pièce jointe
        outlook_service.create_email(
            to="attachment@example.com",
            subject="Test pièces jointes",
            body="Email avec pièces jointes de test.",
        )

        # Créer plusieurs fichiers de test
        test_files = []
        for i in range(3):
            file_path = temp_dir / f"test_file_{i}.txt"
            file_path.write_text(f"Contenu du fichier de test {i}")
            test_files.append(str(file_path))

        # Ajouter les pièces jointes
        for file_path in test_files:
            result = outlook_service.add_attachment(file_path)
            assert result["success"]

        # Sauvegarder en brouillon
        result = outlook_service.save_as_draft()
        assert result["success"]

    def test_contact_management(self, outlook_service):
        """Test : Gestion complète des contacts."""
        # Créer un contact
        result = outlook_service.create_contact(
            first_name="Jean",
            last_name="Dupont",
            email="jean.dupont@example.com",
            company="Société Test",
            phone="+33123456789",
        )
        assert result["success"]

        # Rechercher le contact
        result = outlook_service.search_contacts("Dupont")
        assert result["success"]

        # Modifier le contact
        result = outlook_service.update_contact(
            email="jean.dupont@example.com", phone="+33987654321"
        )
        assert result["success"]

    def test_recurring_appointment(self, outlook_service):
        """Test : Rendez-vous récurrent."""
        start_time = datetime.now() + timedelta(days=1)
        end_time = start_time + timedelta(hours=1)

        result = outlook_service.create_recurring_appointment(
            subject="Réunion hebdomadaire",
            start_time=start_time.isoformat(),
            end_time=end_time.isoformat(),
            recurrence_pattern="weekly",
            recurrence_interval=1,
            recurrence_duration=4,  # 4 semaines
        )
        assert result["success"]


@pytest.mark.slow
class TestOutlookAdvancedIntegration:
    """Tests avancés nécessitant plus de temps."""

    @pytest.fixture(scope="class")
    def outlook_service(self):
        """Initialise le service Outlook."""
        service = OutlookService()
        service.initialize()
        yield service
        service.cleanup()

    @pytest.fixture
    def temp_dir(self):
        """Crée un dossier temporaire."""
        temp_path = Path(tempfile.mkdtemp(prefix="outlook_advanced_"))
        yield temp_path
        # Nettoyage
        for file in temp_path.glob("*"):
            with contextlib.suppress(Exception):
                file.unlink()
        with contextlib.suppress(Exception):
            temp_path.rmdir()

    def test_complex_email_workflow(self, outlook_service, temp_dir):
        """Test : Workflow complexe d'email."""
        # Créer un email avec HTML
        result = outlook_service.create_email(
            to="complex@example.com",
            subject="Email complexe avec HTML",
            body="""
            <html>
            <body>
            <h1>Test d'intégration</h1>
            <p>Ceci est un email avec <strong>formatage HTML</strong>.</p>
            <ul>
            <li>Point 1</li>
            <li>Point 2</li>
            <li>Point 3</li>
            </ul>
            </body>
            </html>
            """,
            html_body=True,
        )
        assert result["success"]

        # Ajouter plusieurs pièces jointes
        attachments = []
        for i in range(3):
            file_path = temp_dir / f"attachment_{i}.txt"
            file_path.write_text(f"Contenu de la pièce jointe {i}")
            attachments.append(str(file_path))

        for attachment in attachments:
            result = outlook_service.add_attachment(attachment)
            assert result["success"]

        # Ajouter des destinataires CC et BCC
        result = outlook_service.add_cc("cc@example.com")
        assert result["success"]

        result = outlook_service.add_bcc("bcc@example.com")
        assert result["success"]

        # Sauvegarder en brouillon
        result = outlook_service.save_as_draft()
        assert result["success"]

    def test_calendar_workflow(self, outlook_service):
        """Test : Workflow complexe de calendrier."""
        # Créer plusieurs rendez-vous
        base_time = datetime.now() + timedelta(days=1)

        appointments = []
        for i in range(3):
            start_time = base_time + timedelta(hours=i * 2)
            end_time = start_time + timedelta(hours=1)

            result = outlook_service.create_appointment(
                subject=f"Rendez-vous {i + 1}",
                start_time=start_time.isoformat(),
                end_time=end_time.isoformat(),
                location=f"Salle {i + 1}",
                body=f"Description du rendez-vous {i + 1}",
            )
            assert result["success"]
            appointments.append(result)

        # Vérifier qu'ils existent
        date = base_time.date()
        result = outlook_service.get_appointments_for_date(date.isoformat())
        assert result["success"]
        assert len(result["appointments"]) >= 3

    def test_contact_batch_operations(self, outlook_service):
        """Test : Opérations par lots sur les contacts."""
        # Créer plusieurs contacts
        contacts = [
            {"first_name": "Alice", "last_name": "Martin", "email": "alice.martin@example.com"},
            {"first_name": "Bob", "last_name": "Durand", "email": "bob.durand@example.com"},
            {"first_name": "Charlie", "last_name": "Leroy", "email": "charlie.leroy@example.com"},
        ]

        for contact in contacts:
            result = outlook_service.create_contact(**contact)
            assert result["success"]

        # Rechercher tous les contacts
        result = outlook_service.search_contacts("")
        assert result["success"]
        assert len(result["contacts"]) >= 3

    def test_performance_large_operations(self, outlook_service):
        """Test : Performance sur opérations volumineuses."""
        import time

        # Test de création rapide de plusieurs éléments
        start_time = time.time()

        # Créer 10 emails en brouillon
        for i in range(10):
            outlook_service.create_email(
                to=f"performance{i}@example.com",
                subject=f"Test performance {i}",
                body=f"Email de test performance {i}",
            )
            outlook_service.save_as_draft()

        elapsed = time.time() - start_time
        print(f"\nTemps pour 10 emails : {elapsed:.2f}s")

        # Les opérations devraient être rapides
        assert elapsed < 30  # Moins de 30 secondes pour 10 emails


def run_integration_tests():
    """Exécute tous les tests d'intégration Outlook."""
    print("=" * 70)
    print("TESTS D'INTÉGRATION OUTLOOK")
    print("=" * 70)
    print()
    print("⚠️  Ces tests vont interagir avec Microsoft Outlook installé.")
    print("Ils vont créer des emails, contacts et rendez-vous réels.")
    print("Assurez-vous qu'Outlook est installé et fermé avant de continuer.")
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
