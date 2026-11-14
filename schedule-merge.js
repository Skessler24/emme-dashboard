/**
 * Three-way merge for schedule workspaces.
 *
 * base   = the workspace both sides last agreed on
 * local  = the copy with the caller's newest edits
 * remote = the copy someone else saved (another tab, an admin, or SharePoint)
 *
 * Boards, days, slots and clients are matched by id. A field that only one side
 * changed keeps that change. If both sides changed the same field, local wins.
 * A row deleted on one side stays deleted unless the other side edited it.
 * Rows added on either side are kept.
 *
 * This file is copied to ../../schedule-merge.js for the browser.
 * Keep the two identical (see tests/check-merge-copy.js).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ScheduleMerge = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function isObject(value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value);
  }

  function equal(a, b) {
    if (a === b) return true;
    if (a === undefined || b === undefined || a === null || b === null) {
      return (a === undefined || a === null || a === '') && (b === undefined || b === null || b === '');
    }
    if (Array.isArray(a) || Array.isArray(b)) {
      if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
      for (let i = 0; i < a.length; i++) if (!equal(a[i], b[i])) return false;
      return true;
    }
    if (isObject(a) && isObject(b)) {
      const keys = Object.keys(a).concat(Object.keys(b));
      for (let i = 0; i < keys.length; i++) if (!equal(a[keys[i]], b[keys[i]])) return false;
      return true;
    }
    return false;
  }

  function clone(value) {
    return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
  }

  function byId(list) {
    const map = {};
    (Array.isArray(list) ? list : []).forEach(function (item) {
      if (item && item.id != null) map[item.id] = item;
    });
    return map;
  }

  function ids(list) {
    return (Array.isArray(list) ? list : []).filter(function (item) {
      return item && item.id != null;
    }).map(function (item) { return item.id; });
  }

  function sameOrder(a, b) {
    const setB = {};
    b.forEach(function (id) { setB[id] = true; });
    const common = a.filter(function (id) { return setB[id]; });
    const setA = {};
    a.forEach(function (id) { setA[id] = true; });
    const other = b.filter(function (id) { return setA[id]; });
    return equal(common, other);
  }

  function insertAfterPredecessor(result, id, order) {
    const at = order.indexOf(id);
    const hasSuccessor = order.slice(at + 1).some(function (other) { return result.indexOf(other) !== -1; });
    if (!hasSuccessor) {
      result.push(id);
      return;
    }
    for (let i = at - 1; i >= 0; i--) {
      const pos = result.indexOf(order[i]);
      if (pos !== -1) {
        result.splice(pos + 1, 0, id);
        return;
      }
    }
    result.unshift(id);
  }

  function mergeList(base, local, remote, mergeItem) {
    const baseMap = byId(base);
    const localMap = byId(local);
    const remoteMap = byId(remote);
    const baseIds = ids(base);
    const localIds = ids(local);
    const remoteIds = ids(remote);

    const keep = {};
    const items = {};
    remoteIds.forEach(function (id) {
      const r = remoteMap[id];
      const l = localMap[id];
      const b = baseMap[id];
      if (l) {
        items[id] = mergeItem(b, l, r);
        keep[id] = true;
      } else if (!b) {
        items[id] = clone(r);
        keep[id] = true;
      } else if (!equal(r, b)) {
        items[id] = clone(r);
        keep[id] = true;
      }
    });
    localIds.forEach(function (id) {
      if (keep[id] || remoteMap[id]) return;
      const l = localMap[id];
      const b = baseMap[id];
      if (!b || !equal(l, b)) {
        items[id] = clone(l);
        keep[id] = true;
      }
    });

    const localLeads = sameOrder(baseIds, remoteIds) && !sameOrder(baseIds, localIds);
    const first = localLeads ? localIds : remoteIds;
    const second = localLeads ? remoteIds : localIds;
    const result = [];
    first.forEach(function (id) { if (keep[id] && result.indexOf(id) === -1) result.push(id); });
    second.forEach(function (id) {
      if (keep[id] && result.indexOf(id) === -1) insertAfterPredecessor(result, id, second);
    });
    return result.map(function (id) { return items[id]; });
  }

  function mergeValue(b, l, r) {
    if (equal(l, b)) return clone(r);
    if (equal(r, b)) return clone(l);
    return clone(l);
  }

  function mergeFields(b, l, r, listFields) {
    const out = {};
    const keys = {};
    [b, l, r].forEach(function (obj) {
      Object.keys(obj || {}).forEach(function (key) { keys[key] = true; });
    });
    Object.keys(keys).forEach(function (key) {
      if (listFields && listFields[key]) {
        out[key] = listFields[key]((b || {})[key], (l || {})[key], (r || {})[key]);
        return;
      }
      const value = mergeValue((b || {})[key], (l || {})[key], (r || {})[key]);
      if (value !== undefined) out[key] = value;
    });
    return out;
  }

  function mergeSlot(b, l, r) {
    return mergeFields(b, l, r);
  }

  function mergeDay(b, l, r) {
    return mergeFields(b, l, r, {
      slots: function (bs, ls, rs) { return mergeList(bs, ls, rs, mergeSlot); }
    });
  }

  function mergeBoard(b, l, r) {
    return mergeFields(b, l, r, {
      days: function (bd, ld, rd) { return mergeList(bd, ld, rd, mergeDay); }
    });
  }

  function mergeClient(b, l, r) {
    return mergeFields(b, l, r);
  }

  function mergeSet(b, l, r) {
    const base = Array.isArray(b) ? b : [];
    const local = Array.isArray(l) ? l : [];
    const remote = Array.isArray(r) ? r : [];
    const out = [];
    function add(value) {
      if (out.indexOf(value) === -1) out.push(value);
    }
    remote.forEach(function (value) {
      if (local.indexOf(value) !== -1 || base.indexOf(value) === -1) add(value);
    });
    local.forEach(function (value) {
      if (base.indexOf(value) === -1) add(value);
    });
    return out;
  }

  function mergeWorkspaces(base, local, remote) {
    const b = base || {};
    const l = local || {};
    const r = remote || {};
    const out = mergeFields(b, l, r, {
      boards: function (bb, lb, rb) { return mergeList(bb, lb, rb, mergeBoard); },
      clients: function (bc, lc, rc) { return mergeList(bc, lc, rc, mergeClient); },
      dxCodes: mergeSet
    });
    out.boards = out.boards || [];
    out.clients = out.clients || [];
    out.dxCodes = out.dxCodes || [];
    return out;
  }

  return { merge: mergeWorkspaces, equal: equal, clone: clone };
}));
