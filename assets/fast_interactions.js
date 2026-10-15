/**
 * Fast Client-Side Interactions for DMB Performance Dashboard
 * ==========================================================
 * Provides instant 0ms UI responsiveness, smooth card clicks,
 * reliable escape key dismissal, and seamless tab switching recovery.
 */

(function () {
    'use strict';

    function initInteractions() {
        // ----------------------------------------------------
        // 1. Navigation Tabs
        // ----------------------------------------------------
        var tabMpr = document.getElementById('nav-tab-mpr');
        var tabDmb = document.getElementById('nav-tab-dmb');

        function setActiveTab(activeName) {
            if (activeName === 'mpr') {
                if (tabMpr) tabMpr.classList.add('navigation-tab-active');
                if (tabDmb) tabDmb.classList.remove('navigation-tab-active');
            } else if (activeName === 'dmb') {
                if (tabDmb) tabDmb.classList.add('navigation-tab-active');
                if (tabMpr) tabMpr.classList.remove('navigation-tab-active');
            }
        }

        if (tabMpr && !tabMpr.hasAttribute('data-bound')) {
            tabMpr.setAttribute('data-bound', 'true');
            tabMpr.addEventListener('click', function (e) {
                e.preventDefault();
                setActiveTab('mpr');
                var mprTarget = document.getElementById('executive-insights-section') || document.getElementById('mpr-section');
                if (mprTarget) {
                    mprTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
                } else {
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }
            });
        }

        if (tabDmb && !tabDmb.hasAttribute('data-bound')) {
            tabDmb.setAttribute('data-bound', 'true');
            tabDmb.addEventListener('click', function (e) {
                e.preventDefault();
                setActiveTab('dmb');
                var dmbTarget = document.getElementById('dmb-section');
                if (dmbTarget) {
                    dmbTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        }

        // ----------------------------------------------------
        // 2. Instant Smooth Visual Feedback on Clickable Cards
        // ----------------------------------------------------
        var rcaCards = document.querySelectorAll('.continuous-red-panel-active');
        rcaCards.forEach(function (card) {
            if (!card.hasAttribute('data-click-bound')) {
                card.setAttribute('data-click-bound', 'true');
                card.addEventListener('click', function () {
                    card.classList.add('card-clicking');
                    setTimeout(function () {
                        card.classList.remove('card-clicking');
                    }, 400);
                });
            }
        });

        var gaugeCards = document.querySelectorAll('.mini-gauge-container-clickable');
        gaugeCards.forEach(function (card) {
            if (!card.hasAttribute('data-click-bound')) {
                card.setAttribute('data-click-bound', 'true');
                card.addEventListener('click', function () {
                    card.classList.add('card-clicking');
                    setTimeout(function () {
                        card.classList.remove('card-clicking');
                    }, 400);
                });
            }
        });

        // ----------------------------------------------------
        // 3. Modal Scroll Lock Sync
        // ----------------------------------------------------
        var rcaModal = document.getElementById('continuous-red-modal');
        var gaugeModal = document.getElementById('function-gauge-modal');
        var mprModal = document.getElementById('mpr-breakdown-modal');
        var isRcaOpen = rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden');
        var isGaugeOpen = gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden');
        var isMprOpen = mprModal && !mprModal.classList.contains('mpr-breakdown-modal-hidden');

        if (isRcaOpen || isGaugeOpen || isMprOpen) {
            document.body.classList.add('modal-open');
            document.documentElement.classList.add('modal-open');
        } else {
            document.body.classList.remove('modal-open');
            document.documentElement.classList.remove('modal-open');
        }
    }

    // ----------------------------------------------------
    // 4. Global Event Delegation: Smooth Scroll on Cards with data-scroll-target
    // ----------------------------------------------------
    document.addEventListener('click', function (e) {
        var card = e.target && e.target.closest ? e.target.closest('[data-scroll-target]') : null;
        if (card) {
            if (e.target && (e.target.tagName === 'SELECT' || e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON')) {
                return;
            }
            var targetId = card.getAttribute('data-scroll-target');
            if (targetId) {
                var targetElem = document.getElementById(targetId);
                if (targetElem) {
                    targetElem.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }
        }
    });

    // ----------------------------------------------------
    // 5. Clean Escape Key Handling (Dispatches Dash native close)
    // ----------------------------------------------------
    if (!window.__modalEscapeBound) {
        window.__modalEscapeBound = true;
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' || e.keyCode === 27) {
                var rcaModal = document.getElementById('continuous-red-modal');
                if (rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden')) {
                    var rcaClose = document.getElementById('close-continuous-red-modal');
                    if (rcaClose) rcaClose.click();
                }

                var gaugeModal = document.getElementById('function-gauge-modal');
                if (gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden')) {
                    var gaugeClose = document.getElementById('close-function-gauge-modal');
                    if (gaugeClose) gaugeClose.click();
                }

                var mprModal = document.getElementById('mpr-breakdown-modal');
                if (mprModal && !mprModal.classList.contains('mpr-breakdown-modal-hidden')) {
                    var mprClose = document.getElementById('close-mpr-breakdown-modal');
                    if (mprClose) mprClose.click();
                }
            }
        });
    }

    // ----------------------------------------------------
    // 6. Tab Switching & Browser Focus Recovery
    // ----------------------------------------------------
    function handleVisibilityRecovery() {
        document.querySelectorAll('.card-clicking, .rca-card-opening').forEach(function (el) {
            el.classList.remove('card-clicking', 'rca-card-opening');
        });
        initInteractions();
    }

    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') {
            handleVisibilityRecovery();
        }
    });

    window.addEventListener('focus', function () {
        handleVisibilityRecovery();
    });

    // ----------------------------------------------------
    // 7. DOM Initialization & Dynamic Component Observation
    // ----------------------------------------------------
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initInteractions);
    } else {
        initInteractions();
    }

    var observer = new MutationObserver(function () {
        initInteractions();
    });
    observer.observe(document.body, { childList: true, subtree: true });
})();
