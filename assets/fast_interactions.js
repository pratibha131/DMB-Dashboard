/**
 * Fast Client-Side Interactions for DMB Performance Dashboard
 * ==========================================================
 * Provides instant 0ms pre-loaded RCA modal opening/closing,
 * smooth delegated card clicks, quick-jump KPI scrolling,
 * reliable Escape key dismissal, and tab/focus state recovery.
 */

(function () {
    'use strict';

    if (window.__dmbFastInteractionsInitialized) {
        return;
    }
    window.__dmbFastInteractionsInitialized = true;
    window.__activeRcaFunction = null;

    var lastRcaOpenTs = 0;
    var pointerDownInfo = null;
    var rafScheduled = false;

    function extractRcaFunction(card) {
        if (!card) return null;
        var attrFn = card.getAttribute('data-rca-open-function');
        if (attrFn) return attrFn;
        var rawId = card.getAttribute('id');
        if (rawId && rawId.charAt(0) === '{') {
            try {
                var parsed = JSON.parse(rawId);
                if (parsed && parsed.function) return parsed.function;
            } catch (err) {
                // ignore malformed id
            }
        }
        return null;
    }

    function syncModalScrollLock() {
        var rcaModal = document.getElementById('continuous-red-modal');
        var gaugeModal = document.getElementById('function-gauge-modal');
        var mprModal = document.getElementById('mpr-breakdown-modal');

        var isRcaOpen = rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden');
        var isGaugeOpen = gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden');
        var isMprOpen = mprModal && !mprModal.classList.contains('mpr-breakdown-modal-hidden');

        if (isRcaOpen || isGaugeOpen || isMprOpen) {
            if (document.body && !document.body.classList.contains('modal-open')) {
                document.body.classList.add('modal-open');
            }
            if (!document.documentElement.classList.contains('modal-open')) {
                document.documentElement.classList.add('modal-open');
            }
        } else {
            if (document.body && document.body.classList.contains('modal-open')) {
                document.body.classList.remove('modal-open');
            }
            if (document.documentElement.classList.contains('modal-open')) {
                document.documentElement.classList.remove('modal-open');
            }
        }
    }

    function activateRcaPanelDom(functionName, resetScroll) {
        if (!functionName) return false;

        var rcaModal = document.getElementById('continuous-red-modal');
        var rcaTitle = document.getElementById('continuous-red-modal-title');
        var rcaMonth = document.getElementById('continuous-red-modal-month');
        var rcaBody = document.getElementById('continuous-red-modal-body');

        if (!rcaModal || !rcaBody) return false;

        var panels = rcaBody.querySelectorAll('.rca-preloaded-panel');
        if (!panels || panels.length === 0) return false;

        var targetPanel = null;
        for (var i = 0; i < panels.length; i++) {
            if (panels[i].getAttribute('data-rca-function') === functionName) {
                targetPanel = panels[i];
                break;
            }
        }

        if (!targetPanel) return false;

        for (var j = 0; j < panels.length; j++) {
            if (panels[j] === targetPanel) {
                if (!panels[j].classList.contains('rca-preloaded-panel-active')) {
                    panels[j].classList.add('rca-preloaded-panel-active');
                }
            } else if (panels[j].classList.contains('rca-preloaded-panel-active')) {
                panels[j].classList.remove('rca-preloaded-panel-active');
            }
        }

        var panelTitle = targetPanel.getAttribute('data-rca-title') || (functionName + ' — Root Cause Analysis');
        var panelMonth = targetPanel.getAttribute('data-rca-month') || '';

        if (rcaTitle && rcaTitle.textContent !== panelTitle) {
            rcaTitle.textContent = panelTitle;
        }
        if (rcaMonth && rcaMonth.textContent !== panelMonth) {
            rcaMonth.textContent = panelMonth;
        }

        window.__activeRcaFunction = functionName;

        if (rcaModal.classList.contains('continuous-red-modal-hidden')) {
            rcaModal.classList.remove('continuous-red-modal-hidden');
        }

        if (resetScroll) {
            rcaBody.scrollTop = 0;
        }

        syncModalScrollLock();
        return true;
    }

    function openRcaModal(functionName, cardEl) {
        if (!functionName) return false;
        var now = Date.now();
        if (now - lastRcaOpenTs < 80 && window.__activeRcaFunction === functionName) {
            return true;
        }
        lastRcaOpenTs = now;

        if (cardEl) {
            cardEl.classList.add('card-clicking');
            setTimeout(function () {
                cardEl.classList.remove('card-clicking');
            }, 180);
        }

        return activateRcaPanelDom(functionName, true);
    }

    function closeRcaModal() {
        var rcaModal = document.getElementById('continuous-red-modal');
        if (rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden')) {
            rcaModal.classList.add('continuous-red-modal-hidden');
        }
        window.__activeRcaFunction = null;
        syncModalScrollLock();
    }

    function syncAfterDomUpdate() {
        rafScheduled = false;
        if (window.__activeRcaFunction) {
            var rcaModal = document.getElementById('continuous-red-modal');
            if (rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden')) {
                activateRcaPanelDom(window.__activeRcaFunction, false);
            } else {
                window.__activeRcaFunction = null;
            }
        }
        syncModalScrollLock();
    }

    function scheduleSync() {
        if (!rafScheduled) {
            rafScheduled = true;
            window.requestAnimationFrame(syncAfterDomUpdate);
        }
    }

    // ----------------------------------------------------
    // 1. Micro-slip Tolerant Pointer Tracking on RCA Cards
    // ----------------------------------------------------
    document.addEventListener('pointerdown', function (e) {
        if (e.button !== undefined && e.button !== 0) return;
        var rcaCard = e.target && e.target.closest ? e.target.closest('.continuous-red-panel-active') : null;
        if (rcaCard && !rcaCard.disabled) {
            var fn = extractRcaFunction(rcaCard);
            if (fn) {
                pointerDownInfo = {
                    fn: fn,
                    card: rcaCard,
                    x: e.clientX,
                    y: e.clientY,
                    ts: Date.now()
                };
                return;
            }
        }
        pointerDownInfo = null;
    }, true);

    document.addEventListener('pointerup', function (e) {
        if (!pointerDownInfo) return;
        var info = pointerDownInfo;
        pointerDownInfo = null;

        if (e.button !== undefined && e.button !== 0) return;
        var dt = Date.now() - info.ts;
        var dx = Math.abs(e.clientX - info.x);
        var dy = Math.abs(e.clientY - info.y);

        // If user pressed on an active RCA card and released with minimal movement,
        // ensure it opens even if the pointer slipped a few pixels off the button edge
        // or the card re-rendered between pointerdown and pointerup.
        var upRcaCard = e.target && e.target.closest ? e.target.closest('.continuous-red-panel-active') : null;
        if (!upRcaCard && dt < 450 && dx < 24 && dy < 24) {
            openRcaModal(info.fn, info.card);
        }
    }, true);

    // ----------------------------------------------------
    // 2. Global Capture-Phase Click Delegation (0ms Instant)
    // ----------------------------------------------------
    document.addEventListener('click', function (e) {
        var target = e.target;
        if (!target || !target.closest) return;

        // A. Open RCA Modal instantly from Preloaded DOM
        var rcaCard = target.closest('.continuous-red-panel-active');
        if (rcaCard && !rcaCard.disabled) {
            var fn = extractRcaFunction(rcaCard);
            if (fn && openRcaModal(fn, rcaCard)) {
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        // B. Close RCA Modal instantly (Back button or Backdrop)
        var rcaClose = target.closest('#close-continuous-red-modal, #continuous-red-modal-backdrop');
        if (rcaClose) {
            closeRcaModal();
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        // C. Quick Jump to KPI Pill inside RCA Modal
        var kpiPill = target.closest('.modal-rca-kpi-pill');
        if (kpiPill) {
            var targetId = kpiPill.getAttribute('data-rca-kpi-target');
            if (!targetId) {
                var rawPillId = kpiPill.getAttribute('id') || '';
                if (rawPillId.charAt(0) === '{') {
                    try {
                        targetId = JSON.parse(rawPillId).target;
                    } catch (err) {
                        targetId = null;
                    }
                }
            }
            if (targetId) {
                var activePanel = document.querySelector('#continuous-red-modal-body .rca-preloaded-panel-active');
                var kpiEl = (activePanel && activePanel.querySelector('[id="' + targetId + '"]')) || document.getElementById(targetId);
                if (kpiEl) {
                    kpiEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    kpiEl.classList.add('kpi-card-highlighted');
                    setTimeout(function () {
                        kpiEl.classList.remove('kpi-card-highlighted');
                    }, 2500);
                }
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        // D. Instant Visual Feedback on Function Gauge Cards
        var gaugeCard = target.closest('.mini-gauge-container-clickable');
        if (gaugeCard && !gaugeCard.disabled) {
            gaugeCard.classList.add('card-clicking');
            setTimeout(function () {
                gaugeCard.classList.remove('card-clicking');
            }, 220);
            return;
        }

        // E. Navigation Tabs
        var tabMpr = target.closest('#nav-tab-mpr');
        if (tabMpr) {
            e.preventDefault();
            tabMpr.classList.add('navigation-tab-active');
            var otherDmb = document.getElementById('nav-tab-dmb');
            if (otherDmb) otherDmb.classList.remove('navigation-tab-active');
            var mprTarget = document.getElementById('executive-insights-section') || document.getElementById('mpr-section');
            if (mprTarget) {
                mprTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
            } else {
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }
            return;
        }

        var tabDmb = target.closest('#nav-tab-dmb');
        if (tabDmb) {
            e.preventDefault();
            tabDmb.classList.add('navigation-tab-active');
            var otherMpr = document.getElementById('nav-tab-mpr');
            if (otherMpr) otherMpr.classList.remove('navigation-tab-active');
            var dmbTarget = document.getElementById('dmb-section');
            if (dmbTarget) {
                dmbTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            return;
        }

        // F. Smooth Scroll on Summary Cards with data-scroll-target
        var scrollCard = target.closest('[data-scroll-target]');
        if (scrollCard) {
            if (target.tagName === 'SELECT' || target.tagName === 'INPUT' || target.tagName === 'BUTTON') {
                return;
            }
            var scrollTargetId = scrollCard.getAttribute('data-scroll-target');
            if (scrollTargetId) {
                var targetElem = document.getElementById(scrollTargetId);
                if (targetElem) {
                    targetElem.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }
        }
    }, true);

    // ----------------------------------------------------
    // 3. Escape Key Dismissal
    // ----------------------------------------------------
    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' || e.keyCode === 27) {
            var rcaModal = document.getElementById('continuous-red-modal');
            if (rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden')) {
                closeRcaModal();
                return;
            }

            var gaugeModal = document.getElementById('function-gauge-modal');
            if (gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden')) {
                var gaugeClose = document.getElementById('close-function-gauge-modal');
                if (gaugeClose) gaugeClose.click();
                return;
            }

            var mprModal = document.getElementById('mpr-breakdown-modal');
            if (mprModal && !mprModal.classList.contains('mpr-breakdown-modal-hidden')) {
                var mprClose = document.getElementById('close-mpr-breakdown-modal');
                if (mprClose) mprClose.click();
            }
        }
    });

    // ----------------------------------------------------
    // 4. Tab Switching & Browser Focus Recovery
    // ----------------------------------------------------
    function handleVisibilityRecovery() {
        pointerDownInfo = null;
        document.querySelectorAll('.card-clicking, .rca-card-opening').forEach(function (el) {
            el.classList.remove('card-clicking', 'rca-card-opening');
        });
        scheduleSync();
    }

    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') {
            handleVisibilityRecovery();
        }
    });

    window.addEventListener('focus', handleVisibilityRecovery);

    // ----------------------------------------------------
    // 5. Lightweight rAF-Throttled DOM Observer
    // ----------------------------------------------------
    function attachObserver() {
        if (!document.body) {
            window.requestAnimationFrame(attachObserver);
            return;
        }
        syncAfterDomUpdate();
        var observer = new MutationObserver(scheduleSync);
        observer.observe(document.body, {
            childList: true,
            subtree: true,
            attributes: true,
            attributeFilter: ['class']
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachObserver);
    } else {
        attachObserver();
    }
})();
