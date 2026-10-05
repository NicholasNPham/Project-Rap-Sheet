"""Named error types for RAP SHEET.

Every error is either a RowProblem (this row needs a person, the run carries
on) or a SystemProblem (something RAP SHEET depends on is broken, so the run
stops). main.py routes on which family an error belongs to, not on individual
error types.

Ported from PROJECT-DALYN's src/exceptions.py. DocumentProblem is called
RowProblem here because the unit of work is a spreadsheet row, not an email
attachment, and the Graph-specific errors are gone with the mailbox layer.
"""


class RapSheetError(Exception):
    """Base class for every RAP SHEET-specific error."""


class RowProblem(RapSheetError):
    """Something is wrong with this row or its PDF. Flag it and move on."""


class SystemProblem(RapSheetError):
    """CCIS, STAC or RAP SHEET itself is broken. Stop the run."""


class ConfigProblem(SystemProblem):
    """config.yaml is missing something, or holds a combination that is refused."""
