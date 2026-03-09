from django.core.management.base import BaseCommand
from api.models import User
import os

class Command(BaseCommand):
    help = "Create system admin with token"

    def handle(self, *args, **kwargs):
        email = os.getenv("ADMIN_EMAIL", "admin@example.com")

        if User.objects.filter(email=email).exists():
            self.stdout.write("Admin already exists")
            return

        admin = User.objects.create_user(
            email=email,
            is_admin=True
        )

        self.stdout.write(self.style.SUCCESS(
            f"Admin created\nEmail: {admin.email}\nToken: {admin.token}"
        ))