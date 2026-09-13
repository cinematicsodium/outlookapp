try:
    from pywintypes import com_error
except ImportError:
    COM_ERRORS = (AttributeError,)
else:
    COM_ERRORS = (AttributeError, com_error)


class OutlookError(Exception):
    """Raised when Outlook cannot complete an operation."""
