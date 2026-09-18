// Canvas-side clustered force for statement graphs.
// Graphviz compound clusters are not used: the standalone HTML must relayout
// when Cluster by changes, and the system Graphviz binary is optional.
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.SemanticVizLayout = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var MIXED_SHEET = "mixed";
  var TICKS = 55;
  var CLUSTER_ATTRACT = 0.16;
  var PLAIN_CENTER = 0.03;
  var LINK_STRENGTH = 0.06;
  var CHARGE = 160;
  var PAD = 48;

  function clusteredLayoutAllowed(primitiveCount, boxMaxPrimitives) {
    return primitiveCount <= boxMaxPrimitives;
  }

  function clusterKey(n, mode) {
    if (mode === "role") {
      return n.is_remainder ? "remainder" : n.direction || "internal";
    }
    if (mode === "sheet") return n.sheet || MIXED_SHEET;
    if (mode === "series") return String(n.series_id);
    return "";
  }

  function hashKey(key) {
    var h = 0;
    var s = String(key);
    for (var i = 0; i < s.length; i++) {
      h = (Math.imul(h, 31) + s.charCodeAt(i)) | 0;
    }
    return h;
  }

  function visibleNodes(nodes) {
    var out = [];
    for (var i = 0; i < nodes.length; i++) {
      if (!nodes[i]._hidden) out.push(nodes[i]);
    }
    return out;
  }

  function sizeNodes(nodes, compact) {
    nodes.forEach(function (n) {
      n._w = compact ? 6 : Math.max(72, String(n._label || n.id || "").length * 7 + 16);
      n._h = compact ? 6 : 28;
    });
  }

  function groupMembers(nodes, splitByRank) {
    var groups = {};
    visibleNodes(nodes).forEach(function (n) {
      var key = n._cluster;
      if (!key) return;
      var slot = splitByRank ? key + "\0" + (n.rank || 0) : key;
      var g = groups[slot] || (groups[slot] = { key: key, rank: n.rank || 0, members: [] });
      g.members.push(n);
    });
    return Object.keys(groups)
      .sort()
      .map(function (slot) {
        return groups[slot];
      });
  }

  function clusterHulls(nodes, options) {
    var splitByRank = !!(options && options.splitByRank);
    var pad = options && options.pad != null ? options.pad : 10;
    return groupMembers(nodes, splitByRank).map(function (g) {
      var x0 = Infinity;
      var y0 = Infinity;
      var x1 = -Infinity;
      var y1 = -Infinity;
      g.members.forEach(function (n) {
        var hw = (n._w || 6) / 2;
        var hh = (n._h || 6) / 2;
        x0 = Math.min(x0, n._x - hw);
        y0 = Math.min(y0, n._y - hh);
        x1 = Math.max(x1, n._x + hw);
        y1 = Math.max(y1, n._y + hh);
      });
      return {
        key: g.key,
        rank: splitByRank ? g.rank : null,
        members: g.members,
        x0: x0 - pad,
        y0: y0 - pad,
        x1: x1 + pad,
        y1: y1 + pad,
      };
    });
  }

  function uniqueKeys(nodes) {
    var keys = [];
    visibleNodes(nodes).forEach(function (n) {
      var key = n._cluster || "";
      if (keys.indexOf(key) < 0) keys.push(key);
    });
    keys.sort();
    return keys;
  }

  function separateBoxes(vis) {
    var nVis = vis.length;
    for (var iter = 0; iter < 10; iter++) {
      for (var i = 0; i < nVis; i++) {
        var a = vis[i];
        for (var j = i + 1; j < nVis; j++) {
          var b = vis[j];
          var dx = b._x - a._x;
          var dy = b._y - a._y;
          var ox = (a._w + b._w) / 2 + 10 - Math.abs(dx);
          var oy = (a._h + b._h) / 2 + 8 - Math.abs(dy);
          if (ox <= 0 || oy <= 0) continue;
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
    }
  }

  function layoutClustered(nodes, bundles, compact) {
    sizeNodes(nodes, compact);
    var vis = visibleNodes(nodes);
    if (!vis.length) return { width: 920, height: 640 };

    var keys = uniqueKeys(vis);
    var grouped = keys.length > 1 || (keys.length === 1 && keys[0] !== "");
    var counts = {};
    vis.forEach(function (n) {
      var k = n._cluster || "";
      counts[k] = (counts[k] || 0) + 1;
    });
    var maxN = 1;
    keys.forEach(function (k) {
      if (counts[k] > maxN) maxN = counts[k];
    });
    var cell = Math.max(
      compact ? 140 : 240,
      (compact ? 48 : 86) * Math.sqrt(maxN) + (compact ? 40 : 80)
    );
    var cols = Math.max(1, Math.ceil(Math.sqrt(keys.length)));
    var centroids = {};
    keys.forEach(function (k, i) {
      var col = i % cols;
      var row = Math.floor(i / cols);
      var jitter = hashKey(k);
      centroids[k] = {
        x: PAD + (col + 0.5) * cell + ((jitter % 17) - 8) * 6,
        y: PAD + (row + 0.5) * cell + (((jitter >> 4) % 13) - 6) * 6,
      };
    });

    vis.forEach(function (n) {
      var k = n._cluster || "";
      var c = centroids[k];
      var members = counts[k];
      var local = 0;
      for (var j = 0; j < vis.length; j++) {
        if (vis[j] === n) break;
        if ((vis[j]._cluster || "") === k) local += 1;
      }
      var ang = (2 * Math.PI * local) / Math.max(members, 1);
      var rad = 12 + 10 * Math.sqrt(members);
      n._x = c.x + Math.cos(ang) * rad;
      n._y = c.y + Math.sin(ang) * rad;
    });

    var byId = {};
    var indexOf = {};
    vis.forEach(function (n, i) {
      byId[n.id] = n;
      indexOf[n.id] = i;
    });
    var links = [];
    (bundles || []).forEach(function (b) {
      var a = byId[b.consumer_id];
      var c = byId[b.producer_id];
      if (!a || !c || a === c) return;
      links.push([indexOf[a.id], indexOf[c.id]]);
    });

    var linkDist = compact ? 28 : 96;
    var charge = compact ? 48 : CHARGE;
    var attract = grouped ? CLUSTER_ATTRACT : PLAIN_CENTER;
    var nVis = vis.length;
    for (var tick = 0; tick < TICKS; tick++) {
      var tnorm = TICKS > 1 ? tick / (TICKS - 1) : 1;
      var alpha = Math.pow(1 - tnorm, 0.5);
      var fx = new Array(nVis);
      var fy = new Array(nVis);
      var i;
      for (i = 0; i < nVis; i++) {
        fx[i] = 0;
        fy[i] = 0;
      }
      for (i = 0; i < links.length; i++) {
        var ui = links[i][0];
        var vi = links[i][1];
        var dx = vis[vi]._x - vis[ui]._x;
        var dy = vis[vi]._y - vis[ui]._y;
        var dist = Math.hypot(dx, dy) || 1e-9;
        var fmag = LINK_STRENGTH * (dist - linkDist);
        var ux = (fmag * dx) / dist;
        var uy = (fmag * dy) / dist;
        fx[ui] += ux;
        fy[ui] += uy;
        fx[vi] -= ux;
        fy[vi] -= uy;
      }
      for (i = 0; i < nVis; i++) {
        for (var j = i + 1; j < nVis; j++) {
          dx = vis[j]._x - vis[i]._x;
          dy = vis[j]._y - vis[i]._y;
          var dist2 = dx * dx + dy * dy + 0.01;
          dist = Math.sqrt(dist2);
          var inv = charge / (dist2 * dist);
          fx[i] -= inv * dx;
          fy[i] -= inv * dy;
          fx[j] += inv * dx;
          fy[j] += inv * dy;
        }
      }
      for (i = 0; i < nVis; i++) {
        var n = vis[i];
        var c = centroids[n._cluster || ""];
        fx[i] += attract * (c.x - n._x);
        fy[i] += attract * (c.y - n._y);
        n._x += fx[i] * alpha;
        n._y += fy[i] * alpha;
      }
    }

    separateBoxes(vis);

    var minX = Infinity;
    var minY = Infinity;
    var maxX = -Infinity;
    var maxY = -Infinity;
    vis.forEach(function (n) {
      minX = Math.min(minX, n._x - n._w / 2);
      minY = Math.min(minY, n._y - n._h / 2);
      maxX = Math.max(maxX, n._x + n._w / 2);
      maxY = Math.max(maxY, n._y + n._h / 2);
    });
    var dx0 = PAD - minX;
    var dy0 = PAD - minY;
    vis.forEach(function (n) {
      n._x += dx0;
      n._y += dy0;
    });
    return {
      width: Math.max(920, maxX - minX + 2 * PAD),
      height: Math.max(640, maxY - minY + 2 * PAD),
    };
  }

  return {
    MIXED_SHEET: MIXED_SHEET,
    clusterKey: clusterKey,
    clusteredLayoutAllowed: clusteredLayoutAllowed,
    clusterHulls: clusterHulls,
    layoutClustered: layoutClustered,
  };
});
