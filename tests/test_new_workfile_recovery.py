import unittest
from unittest.mock import MagicMock

from crawler import CCCCrawler


class FakePage:
    def __init__(self):
        self.url = "https://www.cccone.com/workfile-search/workfile"
        self.goto_count = 0
        self.wait_load_count = 0
        self.context = MagicMock()

    def goto(self, *_args, **_kwargs):
        self.goto_count += 1

    def wait_for_load_state(self, *_args, **_kwargs):
        self.wait_load_count += 1

    def wait_for_function(self, *_args, **_kwargs):
        return MagicMock()

    def get_by_text(self, *_args, **_kwargs):
        return MagicMock()

    def wait_for_url(self, *_args, **_kwargs):
        self.url = "https://www.cccone.com/workfile/edit/contacts"

    def locator(self, *_args, **_kwargs):
        locator = MagicMock()
        locator.wait_for.return_value = None
        locator.inner_text.return_value = ""
        return locator


class NewWorkfileRecoveryTests(unittest.TestCase):
    def test_reload_recovery_clicks_control_after_first_empty_list(self):
        page = FakePage()
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.workfiles_url = page.url
        crawler.dismiss_onboarding_guide = MagicMock()
        crawler.dismiss_jumpstart_prompt = MagicMock()
        crawler.capture_failure = MagicMock()
        crawler.checkpoint = MagicMock()
        crawler.sign_in_if_needed = MagicMock()
        crawler.open_workfiles = MagicMock(side_effect=lambda: setattr(crawler, "workfiles_url", page.url))

        control = MagicMock()
        visibility_checks = iter([None, None, control])
        crawler.first_visible = MagicMock(side_effect=lambda _locator: next(visibility_checks))

        crawler.open_new_workfile()

        control.click.assert_called_once_with(timeout=15000, no_wait_after=True)
        self.assertGreaterEqual(page.goto_count, 1)
        self.assertTrue(crawler._workfile_initialized)


if __name__ == "__main__":
    unittest.main()
