/**
 * Fast Client-Side Interactions for DMB Performance Dashboard
 * ==========================================================
 * Provides instant 0ms pre-loaded RCA and Function Gauge modal opening/closing,
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
    window.__activeGaugeFunction = null;

    var lastRcaOpenTs = 0;
    var lastGaugeOpenTs = 0;
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

    function extractGaugeFunction(card) {
        if (!card) return null;
        var attrFn = card.getAttribute('data-gauge-open-function');
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

        var preloaded = window.__dmbModalsPreloaded;
        if (preloaded && preloaded.rca && preloaded.rca[functionName]) {
            var entry = preloaded.rca[functionName];
            var pTitle = entry.title || (functionName + ' — Root Cause Analysis');
            var pMonth = entry.month || '';
            var renderKey = (preloaded.month || '') + '::' + functionName;

            if (rcaBody.getAttribute('data-rendered-key') !== renderKey) {
                rcaBody.innerHTML = entry.html || '';
                rcaBody.setAttribute('data-rendered-key', renderKey);
            }
            if (rcaTitle && rcaTitle.textContent !== pTitle) {
                rcaTitle.textContent = pTitle;
            }
            if (rcaMonth && rcaMonth.textContent !== pMonth) {
                rcaMonth.textContent = pMonth;
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

    function activateGaugePanelDom(functionName, resetScroll) {
        if (!functionName) return false;

        var gaugeModal = document.getElementById('function-gauge-modal');
        var gaugeTitle = document.getElementById('function-gauge-modal-title');
        var gaugeSubtitle = document.getElementById('function-gauge-modal-subtitle');
        var gaugeBody = document.getElementById('function-gauge-modal-body');

        if (!gaugeModal || !gaugeBody) return false;

        var preloaded = window.__dmbModalsPreloaded;
        if (preloaded && preloaded.gauge && preloaded.gauge[functionName]) {
            var entry = preloaded.gauge[functionName];
            var pTitle = entry.title || (functionName + ' — Function KPI Performance Review');
            var renderKey = (preloaded.month || '') + '::' + functionName;

            if (gaugeBody.getAttribute('data-rendered-key') !== renderKey) {
                gaugeBody.innerHTML = entry.html || '';
                gaugeBody.setAttribute('data-rendered-key', renderKey);
            }
            if (gaugeTitle && gaugeTitle.textContent !== pTitle) {
                gaugeTitle.textContent = pTitle;
            }
            if (gaugeSubtitle && gaugeSubtitle.textContent !== '') {
                gaugeSubtitle.textContent = '';
            }

            window.__activeGaugeFunction = functionName;

            if (gaugeModal.classList.contains('function-gauge-modal-hidden')) {
                gaugeModal.classList.remove('function-gauge-modal-hidden');
            }
            if (resetScroll) {
                gaugeBody.scrollTop = 0;
            }
            syncModalScrollLock();
            return true;
        }

        var panels = gaugeBody.querySelectorAll('.gauge-preloaded-panel');
        if (!panels || panels.length === 0) return false;

        var targetPanel = null;
        for (var i = 0; i < panels.length; i++) {
            if (panels[i].getAttribute('data-gauge-function') === functionName) {
                targetPanel = panels[i];
                break;
            }
        }

        if (!targetPanel) return false;

        for (var j = 0; j < panels.length; j++) {
            if (panels[j] === targetPanel) {
                if (!panels[j].classList.contains('gauge-preloaded-panel-active')) {
                    panels[j].classList.add('gauge-preloaded-panel-active');
                }
            } else if (panels[j].classList.contains('gauge-preloaded-panel-active')) {
                panels[j].classList.remove('gauge-preloaded-panel-active');
            }
        }

        var panelTitle = targetPanel.getAttribute('data-gauge-title') || (functionName + ' — Function KPI Performance Review');

        if (gaugeTitle && gaugeTitle.textContent !== panelTitle) {
            gaugeTitle.textContent = panelTitle;
        }
        if (gaugeSubtitle && gaugeSubtitle.textContent !== '') {
            gaugeSubtitle.textContent = '';
        }

        window.__activeGaugeFunction = functionName;

        if (gaugeModal.classList.contains('function-gauge-modal-hidden')) {
            gaugeModal.classList.remove('function-gauge-modal-hidden');
        }

        if (resetScroll) {
            gaugeBody.scrollTop = 0;
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

    function openGaugeModal(functionName, cardEl) {
        if (!functionName) return false;
        var now = Date.now();
        if (now - lastGaugeOpenTs < 80 && window.__activeGaugeFunction === functionName) {
            return true;
        }
        lastGaugeOpenTs = now;

        if (cardEl) {
            cardEl.classList.add('card-clicking');
            setTimeout(function () {
                cardEl.classList.remove('card-clicking');
            }, 180);
        }

        return activateGaugePanelDom(functionName, true);
    }

    function closeRcaModal() {
        var rcaModal = document.getElementById('continuous-red-modal');
        if (rcaModal && !rcaModal.classList.contains('continuous-red-modal-hidden')) {
            rcaModal.classList.add('continuous-red-modal-hidden');
        }
        window.__activeRcaFunction = null;
        syncModalScrollLock();
    }

    function closeGaugeModal() {
        var gaugeModal = document.getElementById('function-gauge-modal');
        if (gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden')) {
            gaugeModal.classList.add('function-gauge-modal-hidden');
        }
        window.__activeGaugeFunction = null;
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
        if (window.__activeGaugeFunction) {
            var gaugeModal = document.getElementById('function-gauge-modal');
            if (gaugeModal && !gaugeModal.classList.contains('function-gauge-modal-hidden')) {
                activateGaugePanelDom(window.__activeGaugeFunction, false);
            } else {
                window.__activeGaugeFunction = null;
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
    window.__dmbSyncOpenModalsFromStore = scheduleSync;

    window.__dmbApplyMonthInstant = function (monthStr) {
        if (!monthStr) return;
        window.__dmbCurrentMonth = monthStr;
        var allStore = window.__dmbAllMonthsStore;
        var monthData = (allStore && allStore[monthStr]) ? allStore[monthStr] : null;
        if (monthData) {
            window.__dmbModalsPreloaded = monthData;
            var cardsContainer = document.getElementById('function-cards-container');
            if (cardsContainer && monthData.cards_html && cardsContainer.getAttribute('data-rendered-month') !== monthStr) {
                cardsContainer.innerHTML = monthData.cards_html;
                cardsContainer.setAttribute('data-rendered-month', monthStr);
            }
            scheduleSync();
        }
    };

    // ----------------------------------------------------
    // 1. Micro-slip Tolerant Pointer Tracking on Cards
    // ----------------------------------------------------
    document.addEventListener('pointerdown', function (e) {
        if (e.button !== undefined && e.button !== 0) return;
        var rcaCard = e.target && e.target.closest ? e.target.closest('.continuous-red-panel-active') : null;
        if (rcaCard && !rcaCard.disabled) {
            var fn = extractRcaFunction(rcaCard);
            if (fn) {
                pointerDownInfo = {
                    type: 'rca',
                    fn: fn,
                    card: rcaCard,
                    x: e.clientX,
                    y: e.clientY,
                    ts: Date.now()
                };
                return;
            }
        }
        var gaugeCard = e.target && e.target.closest ? e.target.closest('.mini-gauge-container-clickable') : null;
        if (gaugeCard && !gaugeCard.disabled) {
            var gFn = extractGaugeFunction(gaugeCard);
            if (gFn) {
                pointerDownInfo = {
                    type: 'gauge',
                    fn: gFn,
                    card: gaugeCard,
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

        if (dt < 500 && dx < 30 && dy < 30) {
            if (info.type === 'rca') {
                openRcaModal(info.fn, info.card);
            } else if (info.type === 'gauge') {
                openGaugeModal(info.fn, info.card);
            }
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

        // C. Open Function Gauge Modal instantly from Preloaded DOM
        var gaugeCard = target.closest('.mini-gauge-container-clickable');
        if (gaugeCard && !gaugeCard.disabled) {
            var gFn = extractGaugeFunction(gaugeCard);
            if (gFn && openGaugeModal(gFn, gaugeCard)) {
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        // D. Close Function Gauge Modal instantly (Back button or Backdrop)
        var gaugeClose = target.closest('#close-function-gauge-modal, #function-gauge-modal-backdrop');
        if (gaugeClose) {
            closeGaugeModal();
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        // E. Quick Jump to KPI Pill inside RCA Modal
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

        // F. Navigation Tabs
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

        // G. Smooth Scroll on Summary Cards with data-scroll-target
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
                closeGaugeModal();
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
    // 5. Scoped Modal-Only DOM Observer (Zero Dropdown/Chart Overhead)
    // ----------------------------------------------------
    var modalsObserved = false;
    function attachModalObservers() {
        if (modalsObserved) return;
        var rcaModal = document.getElementById('continuous-red-modal');
        var rcaBody = document.getElementById('continuous-red-modal-body');
        var gaugeModal = document.getElementById('function-gauge-modal');
        var gaugeBody = document.getElementById('function-gauge-modal-body');
        var mprModal = document.getElementById('mpr-breakdown-modal');

        if (!rcaModal || !rcaBody || !gaugeModal || !gaugeBody || !mprModal) {
            setTimeout(attachModalObservers, 150);
            return;
        }
        modalsObserved = true;
        syncAfterDomUpdate();

        var modalAttrObserver = new MutationObserver(scheduleSync);
        modalAttrObserver.observe(rcaModal, { attributes: true, attributeFilter: ['class'] });
        modalAttrObserver.observe(gaugeModal, { attributes: true, attributeFilter: ['class'] });
        modalAttrObserver.observe(mprModal, { attributes: true, attributeFilter: ['class'] });

        var bodyObserver = new MutationObserver(scheduleSync);
        bodyObserver.observe(rcaBody, { childList: true });
        bodyObserver.observe(gaugeBody, { childList: true });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachModalObservers);
    } else {
        attachModalObservers();
    }
})();
