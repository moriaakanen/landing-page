(function () {
  'use strict';

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  }

  function positionPopover(wrap, popover, options) {
    if (!wrap || !popover) return;
    const opts = options || {};
    const gap = Number(opts.gap || 6);
    const viewportPad = Number(opts.viewportPad || 14);
    const maxHeight = Number(opts.maxHeight || 338);
    const minHeight = Number(opts.minHeight || 180);

    requestAnimationFrame(() => {
      const wrapRect = wrap.getBoundingClientRect();
      const spaceAbove = Math.max(0, wrapRect.top - viewportPad);
      const spaceBelow = Math.max(0, window.innerHeight - wrapRect.bottom - viewportPad);
      const naturalHeight = popover.scrollHeight || popover.getBoundingClientRect().height || maxHeight;
      const preferUp = opts.prefer === 'up';
      const preferDown = opts.prefer === 'down';
      const fitsUp = spaceAbove >= Math.min(naturalHeight, maxHeight) + gap;
      const fitsDown = spaceBelow >= Math.min(naturalHeight, maxHeight) + gap;
      const openUp = preferUp
        ? fitsUp || (!fitsDown && spaceAbove > spaceBelow)
        : preferDown
          ? !(fitsUp && !fitsDown)
          : (fitsUp && (!fitsDown || spaceAbove >= spaceBelow));

      const available = (openUp ? spaceAbove : spaceBelow) - gap;
      const height = clamp(available, minHeight, maxHeight);
      popover.style.maxHeight = `${height}px`;
      popover.style.overflowY = naturalHeight > height ? 'auto' : 'visible';
      popover.style.top = openUp ? 'auto' : `calc(100% + ${gap}px)`;
      popover.style.bottom = openUp ? `calc(100% + ${gap}px)` : 'auto';
      popover.dataset.placement = openUp ? 'up' : 'down';
    });
  }

  window.PakUi = Object.assign({}, window.PakUi, { positionPopover });
})();
