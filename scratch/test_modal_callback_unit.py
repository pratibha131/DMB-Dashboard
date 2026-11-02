import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from unittest.mock import patch, MagicMock
from dash.exceptions import PreventUpdate
import pandas as pd
from app import (
    manage_continuous_red_modal,
    manage_function_gauge_modal,
    default_month,
)

class TestAtomicModalCallbacks(unittest.TestCase):
    def test_rca_modal_open_on_card_click(self):
        # 1. User clicks Quality RCA card
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": '{"type":"continuous-red-card","function":"Quality"}.n_clicks', "value": 1}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "continuous-red-card", "function": "Quality"}
            
            c_class, c_title, c_month, c_body, active_state = manage_continuous_red_modal(
                card_clicks=[1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data=None,
                current_modal_state={"is_open": False, "function": None, "month": "2026-08-01"},
            )
            
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertIn("Quality", c_title)
            self.assertIn("August 2026", c_month)
            self.assertIsNotNone(c_body)
            self.assertTrue(active_state["is_open"])
            self.assertEqual(active_state["function"], "Quality")

    def test_rca_modal_switch_from_marketing_to_quality(self):
        # 2. User was on Marketing, then clicks Quality -> directly receives Quality data atomically
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": '{"type":"continuous-red-card","function":"Quality"}.n_clicks', "value": 1}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "continuous-red-card", "function": "Quality"}
            
            c_class, c_title, c_month, c_body, active_state = manage_continuous_red_modal(
                card_clicks=[1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data=None,
                current_modal_state={"is_open": True, "function": "Marketing", "month": "2026-08-01"},
            )
            
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertIn("Quality", c_title)
            self.assertNotIn("Marketing", c_title)
            self.assertEqual(active_state["function"], "Quality")

    def test_rca_modal_close_on_back_click(self):
        # 3. User clicks Back button
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": "close-continuous-red-modal.n_clicks", "value": 1}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "close-continuous-red-modal"
            
            c_class, c_title, c_month, c_body, active_state = manage_continuous_red_modal(
                card_clicks=[0],
                close_clicks=1,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data=None,
                current_modal_state={"is_open": True, "function": "Quality", "month": "2026-08-01"},
            )
            
            self.assertIn("continuous-red-modal-hidden", c_class)
            self.assertFalse(active_state["is_open"])

    def test_rca_modal_persists_during_live_sync(self):
        # 4. Live sync fires while modal is open -> remains open and refreshes
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": "live-sync-state-store.data", "value": {"ts": 12345}}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "live-sync-state-store"
            
            c_class, c_title, c_month, c_body, active_state = manage_continuous_red_modal(
                card_clicks=[0],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data={"ts": 12345},
                current_modal_state={"is_open": True, "function": "Quality", "month": "2026-08-01"},
            )
            
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertTrue(active_state["is_open"])
            self.assertEqual(active_state["function"], "Quality")

    def test_rca_modal_ignored_during_live_sync_when_closed(self):
        # 5. Live sync fires while modal is closed -> PreventUpdate (no work)
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": "live-sync-state-store.data", "value": {"ts": 12345}}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "live-sync-state-store"
            
            with self.assertRaises(PreventUpdate):
                manage_continuous_red_modal(
                    card_clicks=[0],
                    close_clicks=0,
                    backdrop_clicks=0,
                    month_filter_val="2026-08-01",
                    sync_data={"ts": 12345},
                    current_modal_state={"is_open": False, "function": None, "month": "2026-08-01"},
                )

    def test_gauge_modal_open_on_card_click(self):
        # 6. User clicks Quality Gauge card
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": '{"type":"function-gauge-card","function":"Quality"}.n_clicks', "value": 1}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "function-gauge-card", "function": "Quality"}
            
            g_class, g_title, g_subtitle, g_body, active_state = manage_function_gauge_modal(
                gauge_card_clicks=[1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data=None,
                current_modal_state={"is_open": False, "function": None, "month": "2026-08-01"},
            )
            
            self.assertEqual(g_class, "function-gauge-modal")
            self.assertIn("Quality", g_title)
            self.assertEqual(g_subtitle, "")
            self.assertIsNotNone(g_body)
            self.assertTrue(active_state["is_open"])
            self.assertEqual(active_state["function"], "Quality")

    def test_gauge_modal_close_on_back_click(self):
        # 7. User clicks Back button on Gauge modal
        mock_ctx = MagicMock()
        mock_ctx.triggered = [{"prop_id": "close-function-gauge-modal.n_clicks", "value": 1}]
        
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "close-function-gauge-modal"
            
            g_class, g_title, g_subtitle, g_body, active_state = manage_function_gauge_modal(
                gauge_card_clicks=[0],
                close_clicks=1,
                backdrop_clicks=0,
                month_filter_val="2026-08-01",
                sync_data=None,
                current_modal_state={"is_open": True, "function": "Quality", "month": "2026-08-01"},
            )
            
            self.assertIn("function-gauge-modal-hidden", g_class)
            self.assertFalse(active_state["is_open"])

if __name__ == "__main__":
    unittest.main()
