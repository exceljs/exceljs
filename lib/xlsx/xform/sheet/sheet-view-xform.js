/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');

var VIEW_STATES = {
  frozen: 'frozen',
  frozenSplit: 'frozen',
  split: 'split'
};

var SheetViewXform = module.exports = function() {
};

utils.inherits(SheetViewXform, BaseXform, {
  
  get tag() { return 'sheetView'; },
  
  prepare: function(model) {
    switch(model.state) {
      case 'frozen':
      case 'split':
        break;
      default:
        model.state = 'normal';
        break;
    }
  },

  render: function(xmlStream, model) {
    xmlStream.openNode('sheetView', {
      workbookViewId: model.workbookViewId || 0
    });
    var add = function(name, value, included) {
      if (included) {
        xmlStream.addAttribute(name, value);
      }
    };
    add('tabSelected', '1', model.tabSelected);
    add('showRuler', '0', model.showRuler === false);
    add('showRowColHeaders', '0', model.showRowColHeaders === false);
    add('showGridLines', '0', model.showGridLines === false);
    add('zoomScale', model.zoomScale, model.zoomScale);
    add('zoomScaleNormal', model.zoomScaleNormal, model.zoomScaleNormal);
    add('view', model.style, model.style);
    
    var topLeftCell, xSplit, ySplit, activePane;
    switch(model.state) {
      case 'frozen':
        //  <sheetView tabSelected="1" workbookViewId="0">
        //   <pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>
        //   <selection pane="bottomRight" activeCell="C3" sqref="C3"/>
        //  </sheetView>
        xSplit = model.xSplit || 0;
        ySplit = model.ySplit || 0;
        topLeftCell = model.topLeftCell || colCache.getAddress(ySplit + 1, xSplit + 1).address;
        activePane = (model.xSplit && model.ySplit) ? 'bottomRight' : (model.xSplit ? 'topRight' : 'bottomLeft');
        xmlStream.leafNode('pane', {
          xSplit: model.xSplit || undefined,
          ySplit: model.ySplit || undefined,
          topLeftCell: topLeftCell,
          activePane: activePane,
          state: 'frozen'
        });
        xmlStream.leafNode('selection', {
          pane: activePane,
          activeCell: model.activeCell,
          sqref: model.activeCell
        });
        break;
      case 'split':
        //  <sheetView tabSelected="1" workbookViewId="1">
        //   <pane xSplit="3432" ySplit="1152" topLeftCell="D4" activePane="bottomRight"/>
        //   <selection pane="bottomRight" activeCell="D4" sqref="D4"/>
        //  </sheetView>
        if (model.activePane === 'topLeft') {
          model.activePane = undefined;
        }
        xmlStream.leafNode('pane', {
          xSplit: model.xSplit || undefined,
          ySplit:  model.ySplit || undefined,
          topLeftCell: model.topLeftCell,
          activePane: model.activePane
        });
        xmlStream.leafNode('selection', {
          pane: model.activePane,
          activeCell: model.activeCell,
          sqref: model.activeCell
        });
        break;
      case 'normal':
        if (model.activeCell) {
          xmlStream.leafNode('selection', {
            activeCell: model.activeCell,
            sqref: model.activeCell
          });
        }
        break;
    }
    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    switch(node.name) {
      case 'sheetView':
        this.sheetView = {
          workbookViewId: parseInt(node.attributes.workbookViewId),
          tabSelected: node.attributes.tabSelected === '1',
          showRuler: !(node.attributes.showRuler === '0'),
          showRowColHeaders: !(node.attributes.showRowColHeaders === '0'),
          showGridLines: !(node.attributes.showGridLines === '0'),
          zoomScale: parseInt(node.attributes.zoomScale || 100),
          zoomScaleNormal: parseInt(node.attributes.zoomScaleNormal || 100),
          style: node.attributes.view
        };
        this.pane = undefined;
        this.selections = {};
        return true;
      case 'pane':
        this.pane = {
          xSplit: parseInt(node.attributes.xSplit || 0),
          ySplit: parseInt(node.attributes.ySplit || 0),
          topLeftCell: node.attributes.topLeftCell,
          activePane:  node.attributes.activePane || 'topLeft',
          state: node.attributes.state
        };
        return true;
      case 'selection':
        var name = node.attributes.pane || 'topLeft';
        this.selections[name] = {
          pane: name,
          activeCell: node.attributes.activeCell
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function() {
  },
  parseClose: function(name) {
    var model, selection;
    switch(name) {
      case 'sheetView':
        if (this.sheetView && this.pane) {
          model = this.model = {
            workbookViewId: this.sheetView.workbookViewId,
            state: VIEW_STATES[this.pane.state] || 'split', // split is default
            xSplit: this.pane.xSplit,
            ySplit: this.pane.ySplit,
            topLeftCell: this.pane.topLeftCell,
            showRuler: this.sheetView.showRuler,
            showRowColHeaders: this.sheetView.showRowColHeaders,
            showGridLines: this.sheetView.showGridLines,
            zoomScale: this.sheetView.zoomScale,
            zoomScaleNormal: this.sheetView.zoomScaleNormal
          };
          if (this.model.state === 'split') {
            model.activePane = this.pane.activePane;
          }
          selection = this.selections[this.pane.activePane];
          if (selection && selection.activeCell) {
            model.activeCell = selection.activeCell;
          }
          if (this.sheetView.style) {
            model.style = this.sheetView.style;
          }
        } else {
          model = this.model = {
            workbookViewId: this.sheetView.workbookViewId,
            state: 'normal',
            showRuler: this.sheetView.showRuler,
            showRowColHeaders: this.sheetView.showRowColHeaders,
            showGridLines: this.sheetView.showGridLines,
            zoomScale: this.sheetView.zoomScale,
            zoomScaleNormal: this.sheetView.zoomScaleNormal
          };
          selection = this.selections.topLeft;
          if (selection && selection.activeCell) {
            model.activeCell = selection.activeCell;
          }
          if (this.sheetView.style) {
            model.style = this.sheetView.style;
          }
        }
        return false;
      default:
        return true;
    }
  },
  
  reconcile: function(model, options) {
  }
});
