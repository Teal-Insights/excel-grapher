// Statement-viewer helpers: cluster keys, hulls, and box sizing / overlap
// removal. Force layout itself is shared with the cell viewer (viz_force.js).
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SemanticVizLayout = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var MIXED_SHEET = "mixed";

  function clusterKey(n, mode) {
    if (mode === "role") {
      return n.is_remainder ? "remainder" : n.direction || "internal";
    }
    if (mode === "sheet") return n.sheet || MIXED_SHEET;
    if (mode === "series") return String(n.series_id);
    return "";
  }

  function sizeNodes(nodes, compact) {
    nodes.forEach(function (n) {
      n._w = compact ? 6 : Math.max(72, String(n._label || n.id || "").length * 7 + 16);
      n._h = compact ? 6 : 28;
    });
  }

  function clusterHulls(nodes, options) {
    var pad = options && options.pad != null ? options.pad : 10;
    var groups = {};
    nodes.forEach(function (n) {
      if (n._hidden || !n._cluster) return;
      (groups[n._cluster] || (groups[n._cluster] = [])).push(n);
    });
    return Object.keys(groups)
      .sort()
      .map(function (key) {
        var members = groups[key];
        var x0 = Infinity;
        var y0 = Infinity;
        var x1 = -Infinity;
        var y1 = -Infinity;
        members.forEach(function (n) {
          var hw = (n._w || 6) / 2;
          var hh = (n._h || 6) / 2;
          x0 = Math.min(x0, n._x - hw);
          y0 = Math.min(y0, n._y - hh);
          x1 = Math.max(x1, n._x + hw);
          y1 = Math.max(y1, n._y + hh);
        });
        return {
          key: key,
          members: members,
          x0: x0 - pad,
          y0: y0 - pad,
          x1: x1 + pad,
          y1: y1 + pad,
        };
      });
  }

  // Push overlapping boxes apart along the axis of least overlap.
  function separateBoxes(nodes) {
    var vis = nodes.filter(function (n) {
      return !n._hidden;
    });
    var nVis = vis.length;
    for (var iter = 0; iter < 10; iter++) {
      var moved = false;
      for (var i = 0; i < nVis; i++) {
        var a = vis[i];
        for (var j = i + 1; j < nVis; j++) {
          var b = vis[j];
          var dx = b._x - a._x;
          var dy = b._y - a._y;
          var ox = (a._w + b._w) / 2 + 10 - Math.abs(dx);
          var oy = (a._h + b._h) / 2 + 8 - Math.abs(dy);
          if (ox <= 0 || oy <= 0) continue;
          moved = true;
          if (ox < oy) {
            var sx = dx < 0 ? -0.5 : 0.5;
            a._x -= ox * sx;
            b._x += ox * sx;
          } else {
            var sy = dy < 0 ? -0.5 : 0.5;
            a._y -= oy * sy;
            b._y += oy * sy;
          }
        }
      }
      if (!moved) return;
    }
  }

  return {
    MIXED_SHEET: MIXED_SHEET,
    clusterKey: clusterKey,
    clusterHulls: clusterHulls,
    separateBoxes: separateBoxes,
    sizeNodes: sizeNodes,
  };
});
