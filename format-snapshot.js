/**
 * format-snapshot.js
 * 格式快照系统：通过 WPS Undo 机制支持一键回滚。
 * UndoRecord 不可用时直接跳过（MVP 不做 fallback 逐段快照，太慢）。
 */

/* global window, wps, wpsActionQueue */

(function () {
  'use strict';

  function log(msg) { console.log('[Snapshot] ' + msg); }
  function warn(msg) { console.warn('[Snapshot] ' + msg); }

  window.FormatSnapshot = {
    _undoRecordActive: false,
    _supportsUndoRecord: null,

    _checkUndoRecordSupport: function () {
      if (this._supportsUndoRecord !== null) return this._supportsUndoRecord;
      try {
        var app = wps.WpsApplication();
        var ur = app.UndoRecord;
        this._supportsUndoRecord = !!(ur && typeof ur.StartCustomRecord === 'function');
      } catch (e) {
        this._supportsUndoRecord = false;
      }
      log('UndoRecord support: ' + this._supportsUndoRecord);
      return this._supportsUndoRecord;
    },

    /**
     * 开始 Undo 组 — 直接调用，不走队列
     */
    beginUndoGroup: function (label) {
      if (this._checkUndoRecordSupport()) {
        try {
          var app = wps.WpsApplication();
          app.UndoRecord.StartCustomRecord(label || 'AI 排版操作');
          this._undoRecordActive = true;
          log('UndoRecord started: ' + (label || 'AI 排版操作'));
          return;
        } catch (e) {
          warn('UndoRecord start failed: ' + e.message);
        }
      }
      log('no UndoRecord, skipping snapshot (MVP)');
    },

    /**
     * 结束 Undo 组 — 直接调用，不走队列
     */
    endUndoGroup: function () {
      if (this._undoRecordActive) {
        try {
          var app = wps.WpsApplication();
          app.UndoRecord.EndCustomRecord();
          this._undoRecordActive = false;
          log('UndoRecord ended');
        } catch (e) {
          warn('UndoRecord end failed: ' + e.message);
          this._undoRecordActive = false;
        }
      }
    },

    /**
     * 一键回滚
     */
    rollback: function () {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      if (!doc) throw new Error('没有打开的文档');

      try {
        doc.Undo();
        log('Undo executed');
        return { method: 'undo', success: true };
      } catch (e) {
        warn('Undo failed: ' + e.message);
        throw new Error('撤销失败: ' + e.message);
      }
    },

    canRollback: function () {
      return true; // WPS 自带 Undo 总是可用
    },

    clear: function () {
      this._undoRecordActive = false;
    }
  };
})();
