// Shared in-browser relayout for the cell and statement viewers.
// Mirrors excel_grapher/grapher/viz_layout.py (levels 1 and 2, exact
// repulsion). Settings come from viz_layout_js_config(), injected by the
// Python writer, so both viewers and the precomputed layout stay in step.
// Browser relayout is limited to small graphs or one cluster, so there is no
// grid repulsion here.
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.VizForce = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var GOLDEN_ANGLE = Math.PI * (3 - Math.sqrt(5));

  function tier(size, cfg) {
    if (size <= cfg.tier_small_max) return "small";
    if (size <= cfg.tier_medium_max) return "medium";
    return "large";
  }

  function spiral(count, spacing) {
    var xs = new Float64Array(count);
    var ys = new Float64Array(count);
    for (var k = 0; k < count; k++) {
      var r = spacing * Math.sqrt(k + 0.5);
      xs[k] = r * Math.cos(k * GOLDEN_ANGLE);
      ys[k] = r * Math.sin(k * GOLDEN_ANGLE);
    }
    return { x: xs, y: ys };
  }

  // One force run over all pairs. `o` fields: x, y, edges ([u, v, restLen, k]),
  // charge weights w, anchors ax/ay + anchorK, centerK, targetY + pullK, ticks.
  function simulate(o, cfg) {
    var x = o.x;
    var y = o.y;
    var n = x.length;
    var ticks = o.ticks || cfg.ticks;
    var fx = new Float64Array(n);
    var fy = new Float64Array(n);
    for (var t = 0; t < ticks; t++) {
      var alpha = Math.pow(1 - t / ticks, 0.5);
      fx.fill(0);
      fy.fill(0);
      var i, j, dx, dy;
      for (var e = 0; e < o.edges.length; e++) {
        var ed = o.edges[e];
        dx = x[ed[1]] - x[ed[0]];
        dy = y[ed[1]] - y[ed[0]];
        var dist = Math.max(Math.sqrt(dx * dx + dy * dy), 1e-9);
        var mag = (ed[3] * (dist - ed[2])) / dist;
        fx[ed[0]] += dx * mag;
        fy[ed[0]] += dy * mag;
        fx[ed[1]] -= dx * mag;
        fy[ed[1]] -= dy * mag;
      }
      for (i = 0; i < n; i++) {
        for (j = i + 1; j < n; j++) {
          dx = x[j] - x[i];
          dy = y[j] - y[i];
          var d2 = dx * dx + dy * dy + 0.01;
          var rep = cfg.charge / (d2 * Math.sqrt(d2));
          if (o.w) rep *= o.w[i] * o.w[j];
          fx[i] -= rep * dx;
          fy[i] -= rep * dy;
          fx[j] += rep * dx;
          fy[j] += rep * dy;
        }
      }
      var mx = 0;
      var my = 0;
      if (o.centerK) {
        for (i = 0; i < n; i++) {
          mx += x[i];
          my += y[i];
        }
        mx /= n;
        my /= n;
      }
      for (i = 0; i < n; i++) {
        if (o.ax && o.anchorK) {
          fx[i] += o.anchorK * (o.ax[i] - x[i]);
          fy[i] += o.anchorK * (o.ay[i] - y[i]);
        }
        if (o.centerK) {
          fx[i] -= o.centerK * (x[i] - mx);
          fy[i] -= o.centerK * (y[i] - my);
        }
        if (o.targetY && o.pullK) fy[i] += o.pullK * (o.targetY[i] - y[i]);
        var sx = fx[i] * alpha;
        var sy = fy[i] * alpha;
        var norm = Math.sqrt(sx * sx + sy * sy);
        if (norm > cfg.max_step) {
          sx *= cfg.max_step / norm;
          sy *= cfg.max_step / norm;
        }
        x[i] += sx;
        y[i] += sy;
      }
    }
  }

  function separateDiscs(cx, cy, radii, gap) {
    var k = cx.length;
    for (var iter = 0; iter < 40; iter++) {
      var moved = false;
      for (var a = 0; a < k; a++) {
        for (var b = a + 1; b < k; b++) {
          var dx = cx[b] - cx[a];
          var dy = cy[b] - cy[a];
          var dist = Math.sqrt(dx * dx + dy * dy);
          var overlap = radii[a] + radii[b] + gap - dist;
          if (overlap <= 0) continue;
          moved = true;
          var px = dist < 1e-9 ? 0.5 * gap : (dx * 0.5 * overlap) / dist;
          var py = dist < 1e-9 ? 0 : (dy * 0.5 * overlap) / dist;
          cx[a] -= px;
          cy[a] -= py;
          cx[b] += px;
          cy[b] += py;
        }
      }
      if (!moved) return;
    }
  }

  // Multilevel clustered force. `o`: n, edges ([u, v] pairs), clusters (one
  // label per node), depths (input depth per node), rankPull
  // ("none" | "between" | "everywhere"). Returns {x, y} centered on 0.
  function clusteredLayout(o, cfg) {
    var n = o.n;
    var x = new Float64Array(n);
    var y = new Float64Array(n);
    if (!n) return { x: x, y: y };
    var pull = o.rankPull || "none";
    var depths = pull === "none" ? null : o.depths;
    var ids = {};
    var cid = new Int32Array(n);
    var k = 0;
    for (var i = 0; i < n; i++) {
      var key = String(o.clusters[i]);
      if (!(key in ids)) ids[key] = k++;
      cid[i] = ids[key];
    }
    var sizes = new Float64Array(k);
    for (i = 0; i < n; i++) sizes[cid[i]] += 1;
    var radii = new Float64Array(k);
    var meanR = 0;
    for (var c = 0; c < k; c++) {
      radii[c] = cfg.disc_spacing * Math.sqrt(sizes[c]);
      meanR += radii[c] / k;
    }
    var edges = [];
    for (var e = 0; e < o.edges.length; e++) {
      if (o.edges[e][0] !== o.edges[e][1]) edges.push(o.edges[e]);
    }

    // Level 1: one node per cluster.
    var centers = spiral(k, 2 * meanR + cfg.link_distance);
    if (k > 1) {
      var seen = {};
      var cedges = [];
      edges.forEach(function (ed) {
        var a = cid[ed[0]];
        var b = cid[ed[1]];
        if (a === b) return;
        var lo = Math.min(a, b);
        var hi = Math.max(a, b);
        var pk = lo + ":" + hi;
        if (seen[pk]) return;
        seen[pk] = true;
        cedges.push([lo, hi, radii[lo] + radii[hi] + cfg.link_distance, cfg.link_strength]);
      });
      var w = new Float64Array(k);
      for (c = 0; c < k; c++) w[c] = Math.sqrt(sizes[c]);
      var target = null;
      if (depths) {
        var md = new Float64Array(k);
        for (i = 0; i < n; i++) md[cid[i]] += depths[i] / sizes[cid[i]];
        var lo = Infinity;
        var hi = -Infinity;
        var mean = 0;
        var area = 0;
        for (c = 0; c < k; c++) {
          lo = Math.min(lo, md[c]);
          hi = Math.max(hi, md[c]);
          mean += md[c] / k;
          area += radii[c] * radii[c];
        }
        if (hi > lo) {
          var extent = 3 * Math.sqrt(area);
          target = new Float64Array(k);
          for (c = 0; c < k; c++) target[c] = ((md[c] - mean) / (hi - lo)) * extent;
        }
      }
      simulate(
        {
          x: centers.x,
          y: centers.y,
          edges: cedges,
          w: w,
          centerK: cfg.center_strength,
          targetY: target,
          pullK: cfg.rank_pull,
        },
        cfg
      );
      separateDiscs(centers.x, centers.y, radii, cfg.link_distance);
    }

    // Level 2: every node, anchored to its cluster.
    var maxSize = 0;
    for (c = 0; c < k; c++) maxSize = Math.max(maxSize, sizes[c]);
    var disc = spiral(maxSize, cfg.disc_spacing);
    var slot = new Int32Array(k);
    var ax = new Float64Array(n);
    var ay = new Float64Array(n);
    for (i = 0; i < n; i++) {
      var s = slot[cid[i]]++;
      ax[i] = centers.x[cid[i]];
      ay[i] = centers.y[cid[i]];
      var jitter = (((i * 2654435761) >>> 0) % 1000) / 1000 - 0.5;
      x[i] = ax[i] + disc.x[s] + jitter * 1e-3 * cfg.link_distance;
      y[i] = ay[i] + disc.y[s] - jitter * 1e-3 * cfg.link_distance;
    }
    var nodeEdges = edges.map(function (ed) {
      var same = cid[ed[0]] === cid[ed[1]];
      var kk = same ? cfg.link_strength : cfg.link_strength * cfg.inter_cluster_link_scale;
      return [ed[0], ed[1], cfg.link_distance, kk];
    });
    var targetY = null;
    if (depths && pull === "everywhere") {
      var sum = new Float64Array(k);
      var dlo = new Float64Array(k).fill(Infinity);
      var dhi = new Float64Array(k).fill(-Infinity);
      for (i = 0; i < n; i++) {
        sum[cid[i]] += depths[i];
        dlo[cid[i]] = Math.min(dlo[cid[i]], depths[i]);
        dhi[cid[i]] = Math.max(dhi[cid[i]], depths[i]);
      }
      targetY = new Float64Array(n);
      for (i = 0; i < n; i++) {
        c = cid[i];
        var span = Math.max(dhi[c] - dlo[c], 1);
        targetY[i] = ay[i] + ((depths[i] - sum[c] / sizes[c]) / span) * 2 * radii[c];
      }
    }
    simulate(
      {
        x: x,
        y: y,
        edges: nodeEdges,
        ax: ax,
        ay: ay,
        anchorK: cfg.cluster_attract,
        targetY: targetY,
        pullK: cfg.rank_pull_within,
      },
      cfg
    );
    var cxm = 0;
    var cym = 0;
    for (i = 0; i < n; i++) {
      cxm += x[i] / n;
      cym += y[i] / n;
    }
    for (i = 0; i < n; i++) {
      x[i] -= cxm;
      y[i] -= cym;
    }
    return { x: x, y: y };
  }

  // Relayout the nodes in `ids` in place, as one cluster around their current
  // centroid. `o`: ids, edges ([u, v] over global ids), x, y (global arrays,
  // updated in place), depths, rankPull ("none" keeps no pull).
  function relayoutGroup(o, cfg) {
    var ids = o.ids;
    var local = {};
    ids.forEach(function (id, i) {
      local[id] = i;
    });
    var edges = [];
    o.edges.forEach(function (ed) {
      var a = local[ed[0]];
      var b = local[ed[1]];
      if (a === undefined || b === undefined) return;
      edges.push([a, b]);
    });
    var pull = o.rankPull && o.rankPull !== "none" ? "everywhere" : "none";
    var out = clusteredLayout(
      {
        n: ids.length,
        edges: edges,
        clusters: ids.map(function () {
          return 0;
        }),
        depths: o.depths
          ? ids.map(function (id) {
              return o.depths[id];
            })
          : null,
        rankPull: o.depths ? pull : "none",
      },
      cfg
    );
    var cx = 0;
    var cy = 0;
    ids.forEach(function (id) {
      cx += o.x[id] / ids.length;
      cy += o.y[id] / ids.length;
    });
    ids.forEach(function (id, i) {
      o.x[id] = cx + out.x[i];
      o.y[id] = cy + out.y[i];
    });
  }

  return {
    tier: tier,
    clusteredLayout: clusteredLayout,
    relayoutGroup: relayoutGroup,
  };
});
