"""Read config.yaml, fill in defaults, and refuse combinations that are unsafe.

Ported from PROJECT-DALYN's config_loader. The staged switches are kept
verbatim in spirit, because they are the only thing standing between a test
run and filing a person's criminal history onto a stranger's case:

    upload_enabled: false, save_enabled: false
        Finds the case, checks the name, picks the Type/Subtype, stops.
        Nothing reaches STAC. Start here.

    upload_enabled: true, save_enabled: false
        Also uploads the PDF and proves the Save button is ready, then
        refuses to press it. The file DOES reach STAC's server, so only
        against test STAC.

    upload_enabled: true, save_enabled: true
        Presses Save. Rap sheets go onto cases.

One gate is new, and it is the reason this module refuses rather than warns:
the Type/Subtype pair has no default. PCSO911 hardcoded DISCOVERY/911AUDIO in
three places; guessing a pair here would file every rap sheet in the run under
whatever STAC's matrix dialog happened to list first.
"""

from pathlib import Path
import keyring
import yaml

from exceptions import ConfigProblem

from keyring.errors import KeyringError

_HERE = Path(__file__).resolve().parent

# Works whether this file sits in src/ or flat in the project root. Getting it
# wrong is quiet and confusing rather than loud: with the modules flat, a
# blind .parent.parent lands on the folder ABOVE the project, so the config is
# looked for in a sibling directory and every relative path in it resolves
# against the wrong place.
PROJECT_ROOT = _HERE.parent if _HERE.name == "src" else _HERE

# config/config.yaml is the documented home, a bare config.yaml beside the
# code is accepted too. Checked in that order, and the first that exists wins.
CONFIG_CANDIDATES = (
    PROJECT_ROOT / "config" / "config.yaml",
    PROJECT_ROOT / "config.yaml",
)
CONFIG_PATH = CONFIG_CANDIDATES[0]
EXAMPLE_PATH = PROJECT_ROOT / "config" / "config.example.yaml"


def find_config() -> Path | None:
    """The first config file that exists, or None if neither is there."""
    for candidate in CONFIG_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None

# Values the example file ships with. Reaching a real run with any of these
# still in place means config.yaml was copied and not filled in.
PLACEHOLDERS = frozenset({
    "",
    "https://stac-test.example.com",
    "https://ccis.example.com",
    "Path/to/RAPSHEET.xlsx",
    "FILL-THIS-IN",
})

# What a test instance's url is expected to look like. is_test_instance saying
# true while the url says otherwise is a contradiction worth stopping for: one
# of the two is wrong and the expensive guess is the wrong one.
TEST_URL_MARKERS = ("test", "stage", "staging", "uat", "localhost", "127.0.0.1")

DEFAULTS = {
    "stac": {
        "is_test_instance": True,
        "upload_enabled": False,
        "save_enabled": False,
        "max_attempts": 2,
        "wait_timeout": 10,
        "upload_timeout": 60,
        "fresh_browser": False,
        "action_pause": 0,
        "check_name": True,
    },
    "ccis": {
        "wait_timeout": 5,
        "no_results_wait_timeout": 1,
        "pause_between_actions": 1,
        "headless": False,
    },
    # Defaults match the sheet these cases actually arrive in:
    #   A Case Number | B UCN | C Defendant Name | D DOB
    # with the outcome written to E, the first free column, and row 1 a header.
    # There is no id column, so id_column is 0.
    "excel": {
        "sheet": "Sheet1",
        "case_number_column": 1,
        "ucn_column": 2,
        "name_column": 3,
        "dob_column": 4,
        "outcome_column": 5,
        "id_column": 0,
        "first_data_row": 2,
    },
    "paths": {
        "logs": "logs",
        "pdfs": "pdfs",
        "results": "results",
        "chromedriver": "",
    },
}


def _merge(base: dict, over: dict) -> dict:
    """Overlay `over` onto a copy of `base`, one level into each section."""
    merged = {key: dict(value) if isinstance(value, dict) else value
              for key, value in base.items()}
    for key, value in (over or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key].update(value)
        else:
            merged[key] = value
    return merged


def load_config(path: Path | None = None) -> dict:
    """Read config.yaml, apply defaults, and validate.

    Args:
        path: Override for the config file. Defaults to config/config.yaml.

    Returns:
        The config dict, with paths resolved to absolute Path objects.

    Raises:
        ConfigProblem: The file is missing or unreadable, a required value is
            absent or still a placeholder, or the switches are set to a
            combination that is refused.
    """
    if path:
        path = Path(path)
        if not path.is_file():
            raise ConfigProblem(f"No config file at {path}.")
    else:
        path = find_config()
        if path is None:
            looked = " or ".join(str(c) for c in CONFIG_CANDIDATES)
            raise ConfigProblem(
                f"No config file found. Looked for {looked}. Copy "
                "config.example.yaml to one of those names and fill it in. "
                "config.yaml is gitignored, so it never gets committed."
            )

    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as error:
        raise ConfigProblem(f"Could not parse {path}: {error}") from error

    if not isinstance(raw, dict):
        raise ConfigProblem(f"{path} does not contain a YAML mapping.")

    config = _merge(DEFAULTS, raw)

    # Switches first, because whether the Type/Subtype pair is required
    # depends on upload_enabled.
    _validate_credentials(config)
    _validate_switches(config)
    _validate_type_subtype(config)
    _validate_excel(config)
    _resolve_paths(config)

    return config


def _require(section: dict, key: str, where: str) -> str:
    """Return a non-placeholder string value, or raise."""
    value = str(section.get(key) or "").strip()
    if value in PLACEHOLDERS:
        raise ConfigProblem(f"{where}.{key} is not set in config.yaml.")
    return value


KEYRING_SERVICE = "rap-sheet-{}"


def _validate_credentials(config: dict) -> None:
    """Both systems need a url in config.yaml and a login in Credential Manager.

    The login is read from Windows Credential Manager, never from the file. A
    username or password left in config.yaml is refused rather than ignored,
    so a plaintext login cannot quietly creep back in.
    """
    for system in ("ccis", "stac"):
        section = config.setdefault(system, {})
        section["url"] = _require(section, "url", system).rstrip("/")

        leftovers = [key for key in ("username", "password") if key in section]
        if leftovers:
            raise ConfigProblem(
                f"{system}.{' and '.join(leftovers)} is still in config.yaml. "
                "Logins live in Windows Credential Manager now. Delete those "
                "lines, and run store_secret.py if you have not."
            )

        service = KEYRING_SERVICE.format(system)
        try:
            cred = keyring.get_credential(service, None)
        except KeyringError as error:
            raise ConfigProblem(
                f"Could not read {service} from Windows Credential Manager: {error}"
            ) from error

        if cred is None or not cred.username or not cred.password:
            raise ConfigProblem(
                f"No login stored for {service} in Windows Credential Manager. "
                "Run store_secret.py as the Windows account that runs this tool."
            )

        section["username"] = cred.username
        section["password"] = cred.password


def _validate_type_subtype(config: dict) -> None:
    """The Type/Subtype pair is required to upload, and has no default.

    Every rap sheet in a run is filed under this one pair, so a wrong value is
    not one misfiled document, it is all of them. There is deliberately nothing
    to fall back to.

    Required only when upload_enabled is true. With uploads off the pair is
    genuinely never used: add_documents returns before the matrix dialog is
    ever opened. Demanding it anyway would block the one pass most worth
    running first, which signs in to both systems, resolves every case number
    and checks every name against the sheet. That pass is how the spreadsheet
    gets validated, and it should not wait on a code nobody has looked up yet.
    """
    stac = config["stac"]

    document_type = str(stac.get("document_type") or "").strip()
    subtype = str(stac.get("subtype") or "").strip()

    missing = [
        name for name, value in
        (("document_type", document_type), ("subtype", subtype))
        if value in PLACEHOLDERS
    ]

    if missing and stac["upload_enabled"]:
        raise ConfigProblem(
            f"stac.{' and stac.'.join(missing)} must be set before anything can "
            "be uploaded. Read the pair off STAC's Type/Subtype dialog exactly "
            "as its 'Image Type' and 'Image Sub Type' columns spell it, "
            "including spaces. There is no default on purpose: this one pair is "
            "used for every row in the run.\n\n"
            "To run without it, set stac.upload_enabled to false. That pass "
            "still signs in to both systems, finds every case and checks every "
            "name, which is what validates the spreadsheet."
        )

    # Recorded so main can say so once at startup rather than per row.
    config["stac"]["pair_is_set"] = not missing

    # Upper-cased once here rather than at every comparison. The matrix
    # dialog's own cells are compared upper-cased, so a lowercase config
    # value would never match a row.
    stac["document_type"] = document_type.upper()
    stac["subtype"] = subtype.upper()


def _validate_switches(config: dict) -> None:
    """Refuse switch combinations that cannot mean what they say.

    Three of them:

        save without upload   there would be nothing to save.

        is_test_instance true, url does not look like a test instance.
            One of the two is wrong, and assuming the url is right would
            save to live STAC while the config says test.

        is_test_instance false   a live run. Allowed, but it has to be
            deliberate, so it is logged loudly by main rather than slipping
            through as a default. DEFAULTS has it true.
    """
    stac = config["stac"]

    stac["is_test_instance"] = bool(stac.get("is_test_instance", True))
    stac["upload_enabled"] = bool(stac.get("upload_enabled", False))
    stac["save_enabled"] = bool(stac.get("save_enabled", False))

    if stac["save_enabled"] and not stac["upload_enabled"]:
        raise ConfigProblem(
            "stac.save_enabled is true but stac.upload_enabled is false. "
            "There would be nothing to save. Turn upload on first and watch "
            "a run reach the Save button before enabling save."
        )

    url = stac["url"].lower()
    looks_like_test = any(marker in url for marker in TEST_URL_MARKERS)

    if stac["is_test_instance"] and not looks_like_test:
        raise ConfigProblem(
            f"stac.is_test_instance is true but stac.url is {stac['url']}, "
            "which does not look like a test instance. Either point the url "
            "at test STAC, or set is_test_instance to false if a live run is "
            "genuinely intended."
        )

    try:
        stac["max_attempts"] = max(1, int(stac.get("max_attempts", 2)))
        stac["wait_timeout"] = int(stac.get("wait_timeout", 10))
        stac["upload_timeout"] = int(stac.get("upload_timeout", 60))
        stac["action_pause"] = float(stac.get("action_pause", 0) or 0)
    except (TypeError, ValueError) as error:
        raise ConfigProblem(f"A stac timeout value is not a number: {error}") from error


def _validate_excel(config: dict) -> None:
    """The spreadsheet must exist and its columns must be distinct.

    Two settings pointing at the same column is the quiet version of this
    going wrong: a ucn_column left at its default while the UCN really sits
    in the outcome column would read the previous run's outcome text as a
    case number.
    """
    excel = config["excel"]

    workbook = str(excel.get("file") or "").strip()
    if workbook in PLACEHOLDERS:
        raise ConfigProblem("excel.file is not set in config.yaml.")

    workbook_path = Path(workbook)
    if not workbook_path.is_absolute():
        workbook_path = PROJECT_ROOT / workbook_path
    if not workbook_path.is_file():
        raise ConfigProblem(f"excel.file points at {workbook_path}, which is not a file.")
    excel["file"] = workbook_path

    # id_column and case_number_column may be 0, meaning the sheet has no such
    # column. The other three are what the run cannot work without: a name and
    # a date of birth to search CCIS with, and somewhere to record the outcome.
    REQUIRED = ("ucn_column", "dob_column", "name_column", "outcome_column")
    OPTIONAL = ("id_column", "case_number_column")

    columns = {}
    for key in REQUIRED + OPTIONAL:
        try:
            number = int(excel.get(key) or 0)
        except (TypeError, ValueError) as error:
            raise ConfigProblem(f"excel.{key} is not a number.") from error

        if number < 0:
            raise ConfigProblem(f"excel.{key} cannot be negative.")
        if number == 0:
            if key in REQUIRED:
                raise ConfigProblem(
                    f"excel.{key} must be 1 or greater (column A is 1)."
                )
            # Absent, which is allowed. Not entered into the clash check,
            # since several absent columns are not a collision.
            excel[key] = 0
            continue

        excel[key] = number
        columns.setdefault(number, []).append(key)

    clashes = {
        number: keys for number, keys in columns.items() if len(keys) > 1
    }
    if clashes:
        described = "; ".join(
            f"column {number} is set for {' and '.join(keys)}"
            for number, keys in clashes.items()
        )
        raise ConfigProblem(
            f"Two excel settings point at the same column: {described}. "
            "Each one the sheet has needs its own column. Set id_column or "
            "case_number_column to 0 if the sheet does not have that column."
        )

    try:
        excel["first_data_row"] = max(1, int(excel.get("first_data_row", 1)))
    except (TypeError, ValueError) as error:
        raise ConfigProblem("excel.first_data_row is not a number.") from error


def _resolve_paths(config: dict) -> None:
    """Turn the path settings into absolute Paths under the project root."""
    paths = config["paths"]
    for key in ("logs", "pdfs", "results"):
        value = Path(str(paths.get(key) or key))
        paths[key] = value if value.is_absolute() else PROJECT_ROOT / value

    driver = str(paths.get("chromedriver") or "").strip()
    paths["chromedriver"] = driver or None