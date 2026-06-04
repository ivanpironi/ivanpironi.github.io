(function () {
  var d = new Date();
  var label = d.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  document.querySelectorAll('.footer-date').forEach(function (el) {
    el.textContent = label;
  });
})();

window.renderModelAssumptions = function (a, elId) {
  var el = document.getElementById(elId);
  if (!el || !a) return;
  el.innerHTML =
    'Model assumptions &nbsp;&middot;&nbsp; ' +
    'Retainer conversion ' + a.retainerConvPct + '%' +
    ' &nbsp;&middot;&nbsp; retainer annual TCV = ' + a.retainerTcvPct + '% of initial engagement' +
    ' &nbsp;&middot;&nbsp; retainer begins ' + a.retainerLagMonths + ' months after initial close, runs ' + a.retainerInitMonths + ' months, renews once to M36' +
    ' &nbsp;&middot;&nbsp; Year&#8209;3 accounts at ' + a.y3OrigProbPct + '% origination probability.' +
    ' Methodology consistent with The Numbers.';
};
