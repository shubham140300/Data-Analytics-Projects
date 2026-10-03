from getpass import getpass

from django.conf import settings
from django.contrib.auth.hashers import make_password
from django.core.management.base import BaseCommand, CommandError

from website.models import WorkspaceUser


class Command(BaseCommand):
    help = "Create or reset the local workspace owner username and password."

    def add_arguments(self, parser):
        parser.add_argument(
            "--if-needed", action="store_true",
            help="Only prompt when the configured owner account has not been initialized yet.",
        )

    def handle(self, *args, **options):
        owner_id = settings.WORKSPACE_OWNER_ID
        username = settings.WORKSPACE_OWNER_USERNAME.strip().casefold()
        if not owner_id.isdigit():
            raise CommandError("WORKSPACE_OWNER_ID must contain digits only.")
        if not username:
            raise CommandError("Set WORKSPACE_OWNER_USERNAME to a non-empty username.")
        existing_owner = WorkspaceUser.objects.filter(user_id=owner_id).first()
        if (
            options["if_needed"]
            and existing_owner
            and existing_owner.username == username
            and existing_owner.password_hash
            and existing_owner.is_active
        ):
            self.stdout.write("The workspace owner account is already configured.")
            return
        collision = WorkspaceUser.objects.filter(username__iexact=username).exclude(user_id=owner_id).exists()
        if collision:
            raise CommandError("That owner username is already assigned to another workspace user.")

        self.stdout.write(f"Set a password for owner account {username!r} (owner ID {owner_id}).")
        password = getpass("New password (at least 10 characters): ")
        if len(password) < 10:
            raise CommandError("Passwords must be at least 10 characters long.")
        confirmation = getpass("Confirm password: ")
        if password != confirmation:
            raise CommandError("The passwords did not match. Run the command again.")

        WorkspaceUser.objects.update_or_create(
            user_id=owner_id,
            defaults={
                "username": username,
                "password_hash": make_password(password),
                "must_change_password": False,
                "can_read": True,
                "can_write": True,
                "can_execute": True,
                "is_active": True,
            },
        )
        self.stdout.write(self.style.SUCCESS(f"Owner account {username!r} is ready for password sign-in."))
