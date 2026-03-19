import importlib.util
import queue
import sys
import threading
import types
import unittest
from contextlib import contextmanager
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MISSING = object()


def _module(name, **attrs):
    module = types.ModuleType(name)
    for key, value in attrs.items():
        setattr(module, key, value)
    return module


@contextmanager
def _temporary_modules(modules):
    previous = {}
    try:
        for name, module in modules.items():
            previous[name] = sys.modules.get(name, MISSING)
            sys.modules[name] = module
        yield
    finally:
        for name, module in previous.items():
            if module is MISSING:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = module


def _load_module(module_name, relative_path, stubs):
    spec = importlib.util.spec_from_file_location(module_name, ROOT / relative_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    with _temporary_modules(stubs):
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
    return module


def _load_test_modules():
    package = _module("test_pds_generator")
    package.__path__ = [str(ROOT / "pds_generator")]
    gui_package = _module("test_pds_generator.gui")
    gui_package.__path__ = [str(ROOT / "pds_generator" / "gui")]

    pandas_module = _module("pandas", isna=lambda value: value is None)

    class _RequestException(Exception):
        pass

    requests_module = _module(
        "requests",
        RequestException=_RequestException,
        get=lambda *args, **kwargs: None,
    )

    class _UnidentifiedImageError(Exception):
        pass

    image_module = _module("PIL.Image", open=lambda *args, **kwargs: None, LANCZOS=1)
    imagetk_module = _module("PIL.ImageTk", PhotoImage=object)
    pil_module = _module(
        "PIL",
        Image=image_module,
        ImageTk=imagetk_module,
        UnidentifiedImageError=_UnidentifiedImageError,
    )

    class _FakeTk:
        pass

    class _FakeToplevel:
        pass

    class _FakeCanvas:
        pass

    class _FakeFont:
        def configure(self, **kwargs):
            return None

        def copy(self):
            return self

    tk_module = _module(
        "tkinter",
        Tk=_FakeTk,
        Toplevel=_FakeToplevel,
        Canvas=_FakeCanvas,
        Menu=lambda *args, **kwargs: None,
        TclError=RuntimeError,
        StringVar=object,
        BooleanVar=object,
    )
    ttk_module = _module("tkinter.ttk")
    filedialog_module = _module("tkinter.filedialog")
    messagebox_module = _module("tkinter.messagebox")
    colorchooser_module = _module("tkinter.colorchooser")
    font_module = _module("tkinter.font", nametofont=lambda name: _FakeFont())
    tk_module.ttk = ttk_module
    tk_module.filedialog = filedialog_module
    tk_module.messagebox = messagebox_module
    tk_module.colorchooser = colorchooser_module
    tk_module.font = font_module

    text_layout_module = _module(
        "test_pds_generator.text_layout",
        DEFAULT_FONT_FAMILY="Arial",
        fit_text_lines=lambda *args, **kwargs: [],
        pdf_font_name=lambda bold=False: "Helvetica-Bold" if bold else "Helvetica",
    )

    base_stubs = {
        "pandas": pandas_module,
        "requests": requests_module,
        "PIL": pil_module,
        "PIL.Image": image_module,
        "PIL.ImageTk": imagetk_module,
        "tkinter": tk_module,
        "tkinter.ttk": ttk_module,
        "tkinter.filedialog": filedialog_module,
        "tkinter.messagebox": messagebox_module,
        "tkinter.colorchooser": colorchooser_module,
        "tkinter.font": font_module,
        "test_pds_generator": package,
        "test_pds_generator.gui": gui_package,
        "test_pds_generator.text_layout": text_layout_module,
    }

    elements_module = _load_module(
        "test_pds_generator.elements",
        "pds_generator/elements.py",
        base_stubs,
    )

    groups_module = _module(
        "test_pds_generator.groups",
        GroupArea=type("GroupArea", (), {}),
        GroupEditor=type("GroupEditor", (), {}),
    )
    ui_layout_module = _module("test_pds_generator.gui.ui_layout", setup_ui=lambda app: None)
    tooltips_module = _module(
        "test_pds_generator.gui.tooltips",
        Tooltip=type("Tooltip", (), {"__init__": lambda self, *args, **kwargs: None}),
    )
    pdf_export_module = _module(
        "test_pds_generator.gui.pdf_export",
        generate_pds=lambda app: False,
        draw_pdf_element=lambda *args, **kwargs: None,
    )
    pdf_settings_module = _module(
        "test_pds_generator.pdf_settings",
        DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT=85,
        get_pdf_image_compression_profile=lambda *args, **kwargs: {},
        normalize_pdf_image_compression_percent=lambda value: value,
    )
    config_io_module = _module(
        "test_pds_generator.gui.config_io",
        save_config=lambda app: None,
        load_config=lambda app, startup=False, path=None: None,
    )
    mailer_module = _module(
        "test_pds_generator.gui.mailer",
        normalize_mail_config=lambda config: config,
    )
    windows_auth_module = _module("test_pds_generator.gui.windows_auth")
    excel_tracking_module = _module(
        "test_pds_generator.gui.excel_tracking",
        is_tracking_column=lambda name: False,
    )
    excel_io_module = _module(
        "test_pds_generator.excel_io",
        read_excel_data=lambda path: {},
        detect_formula_columns=lambda path: {},
        collect_formula_samples=lambda path: {},
    )
    locks_module = _module(
        "test_pds_generator.gui.locks",
        acquire_lock=lambda *args, **kwargs: None,
        release_lock=lambda *args, **kwargs: None,
        _lock_path=lambda path: path,
    )
    number_format_module = _module(
        "test_pds_generator.number_format",
        round_numeric_value=lambda value, *args, **kwargs: value,
    )
    github_utils_module = _module(
        "test_pds_generator.github_utils",
        get_repo_info=lambda *args, **kwargs: (None, None, None),
        get_remote_commit_info=lambda *args, **kwargs: (None, None),
        get_remote_version=lambda *args, **kwargs: None,
        pull_updates=lambda *args, **kwargs: False,
        get_version=lambda *args, **kwargs: "0",
    )
    image_index_module = _module("test_pds_generator.image_index")

    gui_stubs = dict(base_stubs)
    gui_stubs.update(
        {
            "test_pds_generator.elements": elements_module,
            "test_pds_generator.groups": groups_module,
            "test_pds_generator.gui.ui_layout": ui_layout_module,
            "test_pds_generator.gui.tooltips": tooltips_module,
            "test_pds_generator.gui.pdf_export": pdf_export_module,
            "test_pds_generator.pdf_settings": pdf_settings_module,
            "test_pds_generator.gui.config_io": config_io_module,
            "test_pds_generator.gui.mailer": mailer_module,
            "test_pds_generator.gui.windows_auth": windows_auth_module,
            "test_pds_generator.gui.excel_tracking": excel_tracking_module,
            "test_pds_generator.excel_io": excel_io_module,
            "test_pds_generator.gui.locks": locks_module,
            "test_pds_generator.number_format": number_format_module,
            "test_pds_generator.github_utils": github_utils_module,
            "test_pds_generator.image_index": image_index_module,
        }
    )

    gui_module = _load_module(
        "test_pds_generator.gui.gui",
        "pds_generator/gui/gui.py",
        gui_stubs,
    )
    return elements_module, gui_module


ELEMENTS_MODULE, GUI_MODULE = _load_test_modules()
DraggableElement = ELEMENTS_MODULE.DraggableElement
PDSGeneratorGUI = GUI_MODULE.PDSGeneratorGUI


class _DummyGUI:
    def __init__(self, main_thread_id):
        self._main_thread_id = main_thread_id
        self._ui_call_queue = queue.Queue()
        self._ui_call_shutdown = False
        self._ui_call_queue_after_id = None
        self.after_calls = []
        self.reported = []
        self.rescheduled = 0

    def after(self, delay_ms, callback):
        self.after_calls.append((delay_ms, callback))
        return "after-id"

    def report_callback_exception(self, exc_type, exc, tb):
        self.reported.append((exc_type, exc, tb))

    def _schedule_ui_call_queue_drain(self):
        self.rescheduled += 1


class _DummyRoot:
    def __init__(self):
        self.calls = []

    def ui_call(self, func, *args, **kwargs):
        self.calls.append((func, args, kwargs))


class _DummyParent:
    def __init__(self, parent):
        self.parent = parent


class GuiUiDispatchTests(unittest.TestCase):
    def test_ui_call_runs_immediately_on_main_thread(self):
        app = _DummyGUI(threading.get_ident())
        calls = []

        PDSGeneratorGUI.ui_call(app, calls.append, "ok")

        self.assertEqual(calls, ["ok"])
        self.assertTrue(app._ui_call_queue.empty())
        self.assertEqual(app.after_calls, [])

    def test_ui_call_from_worker_thread_enqueues_without_touching_tk(self):
        app = _DummyGUI(main_thread_id=-1)
        calls = []

        PDSGeneratorGUI.ui_call(app, calls.append, "queued")

        self.assertEqual(calls, [])
        self.assertEqual(app.after_calls, [])
        self.assertFalse(app._ui_call_queue.empty())

        PDSGeneratorGUI._drain_ui_call_queue(app)

        self.assertEqual(calls, ["queued"])
        self.assertEqual(app.rescheduled, 1)

    def test_dispatch_ui_walks_parent_chain_to_find_dispatcher(self):
        root = _DummyRoot()
        parent = _DummyParent(root)
        element = DraggableElement.__new__(DraggableElement)
        element.parent = parent
        element.name = "Zdjecie"

        result = DraggableElement._dispatch_ui(element, lambda value: value, 123)

        self.assertTrue(result)
        self.assertEqual(len(root.calls), 1)
        func, args, kwargs = root.calls[0]
        self.assertEqual(args, (123,))
        self.assertEqual(kwargs, {})
        self.assertEqual(func(123), 123)


if __name__ == "__main__":
    unittest.main()
