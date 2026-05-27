"""
CECL Credential Manager — Windows Credential Manager integration.

Stores and retrieves the DATABASE_URL securely via Windows Credential Manager
instead of the plaintext .env file.

Usage:
    # Store the credential (interactive prompt)
    python cecl_credentials.py --store

    # Store with value directly
    python cecl_credentials.py --store --url "postgresql://user:pass@localhost:5432/cecl_migration_db"

    # Verify stored credential
    python cecl_credentials.py --verify

    # Remove stored credential
    python cecl_credentials.py --delete

In application code:
    from cecl_credentials import get_database_url
    db_url = get_database_url()
"""
import os
import sys
import argparse

SERVICE_NAME = 'CECL_Migration_DB'
CREDENTIAL_KEY = 'DATABASE_URL'


def _keyring_available():
    try:
        import keyring
        return True
    except ImportError:
        return False


def get_database_url():
    """Retrieve DATABASE_URL: try Windows Credential Manager first, fall back to .env.

    Returns:
        str: The database connection URL.
    """
    # 1) Try Windows Credential Manager via keyring
    if _keyring_available():
        import keyring
        stored = keyring.get_password(SERVICE_NAME, CREDENTIAL_KEY)
        if stored:
            return stored

    # 2) Fall back to environment / .env file
    from dotenv import load_dotenv
    load_dotenv()
    url = os.getenv('DATABASE_URL')
    if url:
        return url

    raise RuntimeError(
        "DATABASE_URL not found. Either:\n"
        "  1) Store it in Windows Credential Manager:  python cecl_credentials.py --store\n"
        "  2) Set it in a .env file:  DATABASE_URL=postgresql://..."
    )


def store_credential(url=None):
    """Store DATABASE_URL in Windows Credential Manager."""
    import keyring

    if not url:
        url = input("Enter DATABASE_URL: ").strip()
        if not url:
            print("No URL provided. Aborted.")
            return False

    # Basic validation
    if not url.startswith(('postgresql://', 'postgres://')):
        print("WARNING: URL does not start with postgresql:// — storing anyway.")

    keyring.set_password(SERVICE_NAME, CREDENTIAL_KEY, url)
    print(f"Credential stored in Windows Credential Manager under '{SERVICE_NAME}'.")
    return True


def verify_credential():
    """Verify that a credential is stored and can connect."""
    import keyring

    stored = keyring.get_password(SERVICE_NAME, CREDENTIAL_KEY)
    if not stored:
        print("No credential found in Windows Credential Manager.")
        print("Run:  python cecl_credentials.py --store")
        return False

    # Mask password in display
    display = stored
    if '@' in stored:
        prefix = stored.split('@')[0]
        suffix = stored.split('@')[1]
        if ':' in prefix:
            parts = prefix.split(':')
            # postgresql://user:PASSWORD@host
            masked = parts[0] + ':' + parts[1] + ':' + '***'
            display = masked + '@' + suffix

    print(f"Credential found: {display}")

    # Try to connect
    try:
        from sqlalchemy import create_engine, text
        eng = create_engine(stored)
        with eng.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("Database connection: OK")
        return True
    except Exception as e:
        print(f"Database connection: FAILED — {e}")
        return False


def delete_credential():
    """Remove the stored credential."""
    import keyring

    try:
        keyring.delete_password(SERVICE_NAME, CREDENTIAL_KEY)
        print(f"Credential removed from Windows Credential Manager.")
    except keyring.errors.PasswordDeleteError:
        print("No credential found to delete.")


def main():
    parser = argparse.ArgumentParser(description="Manage CECL database credentials")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--store', action='store_true', help='Store credential in Windows Credential Manager')
    group.add_argument('--verify', action='store_true', help='Verify stored credential')
    group.add_argument('--delete', action='store_true', help='Delete stored credential')
    parser.add_argument('--url', help='Database URL to store (prompted if omitted)')
    args = parser.parse_args()

    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Run:  pip install keyring")
        sys.exit(1)

    if args.store:
        store_credential(args.url)
    elif args.verify:
        verify_credential()
    elif args.delete:
        delete_credential()


if __name__ == '__main__':
    main()
