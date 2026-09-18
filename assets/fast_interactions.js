/**
 * Fast Client-Side Interactions for DMB Performance Dashboard
 * ==========================================================
 * Provides instant 0ms UI responsiveness:
 * 1. Instant Tab Switching (MPR vs DMB) with smooth scroll & scroll-spy
 * 2. Instant Modal Dismissal (Close button, Backdrop click, Escape key)
 * 3. Instant Visual Feedback on clickables
 */

(function () {
    'use strict';

    function initInteractions() {
        // ----------------------------------------------------
        // 1. Navigation Tabs (MPR vs DMB)
        // ----------------------------------------------------
        var tabMpr = document.getElementById('nav-tab-mpr');
        var tabDmb = document.getElementById('nav-tab-dmb');

        function setActiveTab(activeName) {
            if (!tabMpr || !tabDmb) {
                tabMpr = document.getElementById('nav-tab-mpr');
                tabDmb = document.getElementById('nav-tab-dmb');
            }
            if (activeName === 'mpr') {
                if (tabMpr) tabMpr.classList.add('navigation-tab-active');
                if (tabDmb) tabDmb.classList.remove('navigation-tab-active');
            } else if (activeName === 'dmb') {
                if (tabDmb) tabDmb.classList.add('navigation-tab-active');
                if (tabMpr) tabMpr.classList.remove('navigation-tab-active');
            }
        }

        if (tabMpr) {
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

        if (tabDmb) {
            tabDmb.addEventListener('click', function (e) {
                e.preventDefault();
                setActiveTab('dmb');
                var dmbTarget = document.getElementById('dmb-section');
                if (dmbTarget) {
                    dmbTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        }

        // Scroll spy to automatically highlight MPR vs DMB as user scrolls
        var dmbSection = document.getElementById('dmb-section');
        if (dmbSection && 'IntersectionObserver' in window) {
            var observer = new IntersectionObserver(function (entries) {
                entries.forEach(function (entry) {
                    if (entry.isIntersecting && entry.boundingClientRect.top <= window.innerHeight * 0.4) {
                        setActiveTab('dmb');
                    } else if (window.scrollY < (dmbSection.offsetTop - 200)) {
                        setActiveTab('mpr');
                    }
                });
            }, { threshold: [0.1, 0.3, 0.5] });
            observer.observe(dmbSection);
        }

        // ----------------------------------------------------
        // 2. Instant Modal Close (Backdrop, Button, Escape)
        // ----------------------------------------------------
        function closeModalInstantly() {
            var modal = document.getElementById('continuous-red-modal');
            if (modal && !modal.classList.contains('continuous-red-modal-hidden')) {
                modal.classList.add('continuous-red-modal-hidden');
            }
        }

        var closeBtn = document.getElementById('close-continuous-red-modal');
        var backdrop = document.getElementById('continuous-red-modal-backdrop');

        if (closeBtn) {
            closeBtn.addEventListener('click', closeModalInstantly);
        }
        if (backdrop) {
            backdrop.addEventListener('click', closeModalInstantly);
        }

        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' || e.keyCode === 27) {
                closeModalInstantly();
            }
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initInteractions);
    } else {
        initInteractions();
    }

    // Re-bind after Dash reloads components
    var observer = new MutationObserver(function () {
        var tabMpr = document.getElementById('nav-tab-mpr');
        if (tabMpr && !tabMpr.hasAttribute('data-bound')) {
            tabMpr.setAttribute('data-bound', 'true');
            initInteractions();
        }
    });
    observer.observe(document.body, { childList: true, subtree: true });
})();
