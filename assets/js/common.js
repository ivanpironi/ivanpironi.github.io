(function () {
  var d = new Date();
  var label = d.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  document.querySelectorAll('.footer-date').forEach(function (el) {
    el.textContent = label;
  });
})();
