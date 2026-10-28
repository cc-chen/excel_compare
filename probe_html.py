# -*- coding: utf-8 -*-
"""在真实 Chrome 里跑 HTML 报告的交互（筛选 + 复制表格），把结果写进 DOM 供 --dump-dom 读取。"""
import os
import sys

base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_t")
src = os.path.join(base, "demo_diff.html")
h = open(src, encoding="utf-8").read()

probe = """
<script>
window.__res = {};
function runProbe(){
  var added = document.querySelector('.bar button[data-f="added"]');
  var all = document.querySelector('.bar button[data-f="all"]');
  flt(added, 'added');
  var tot = 0, vis = 0, visSt = [];
  document.querySelectorAll('#tbl tbody tr').forEach(function(tr){
    tot++; if (tr.style.display !== 'none') { vis++; visSt.push(tr.dataset.st); }
  });
  window.__res.total = tot;
  window.__res.visAfterAddedFilter = vis;
  window.__res.visStatuses = visSt.join(',');

  // 桩掉剪贴板，抓取复制的文本
  Object.defineProperty(navigator, 'clipboard', {configurable: true,
    value: {writeText: function(t){ window.__res.copied = t; return Promise.resolve(); }}});
  flt(all, 'all');
  copyTable();

  setTimeout(function(){
    var out = {
      total: window.__res.total,
      visAfterAddedFilter: window.__res.visAfterAddedFilter,
      visStatuses: window.__res.visStatuses,
      copiedLines: (window.__res.copied || '').split('\\r\\n').length,
      copiedHead: (window.__res.copied || '').split('\\r\\n')[0],
      copiedFirstRow: (window.__res.copied || '').split('\\r\\n')[1]
    };
    var pre = document.createElement('pre');
    pre.id = 'PROBE';
    pre.textContent = 'PROBE_JSON=' + JSON.stringify(out);
    document.body.appendChild(pre);
  }, 300);
}
window.addEventListener('load', runProbe);
</script>
"""
h2 = h.replace("</body>", probe + "</body>")
out = os.path.join(base, "harness.html")
open(out, "w", encoding="utf-8").write(h2)
print(out)
