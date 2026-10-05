import getpass
import keyring

for system in ("ccis", "stac"):
    service = f"rap-sheet-{system}"
    username = input(f"{system.upper()} username (blank to skip): ").strip()
    if not username:
        print(f"Skipped {system}.")
        continue
    pw = getpass.getpass(f"{system.upper()} password: ")

    old = keyring.get_credential(service, None)
    if old:
        keyring.delete_password(service, old.username)
    keyring.set_password(service, username, pw)

    cred = keyring.get_credential(service, None)
    ok = cred and cred.username == username and cred.password == pw
    print(f"{service}: {'stored OK' if ok else 'MISMATCH, try again'}")