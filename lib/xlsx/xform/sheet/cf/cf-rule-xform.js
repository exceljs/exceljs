const uuid = require('uuid');
const BaseXform = require('../../base-xform');
const Range = require('../../../../doc/range');

const DatabarXform = require('./databar-xform');
const ExtLstRefXform = require('./ext-lst-ref-xform');
const FormulaXform = require('./formula-xform');
const ColorScaleXform = require('./color-scale-xform');
const IconSetXform = require('./icon-set-xform');

const extIcons = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

const getTextFormula = model => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.operator) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH("${model.text}",${tl})))`;
    case 'containsBlanks':
      return `LEN(TRIM(${tl}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${tl}))>0`;
    case 'containsErrors':
      return `ISERROR(${tl})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${tl}))`;
    default:
      return undefined;
  }
};

const getTimePeriodFormula = model => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.timePeriod) {
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case 'nextWeek':
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case 'yesterday':
      return `FLOOR(${tl},1)=TODAY()-1`;
    case 'today':
      return `FLOOR(${tl},1)=TODAY()`;
    case 'tomorrow':
      return `FLOOR(${tl},1)=TODAY()+1`;
    case 'last7Days':
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case 'lastMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case 'thisMonth':
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case 'nextMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
};

const opType = attr => {
  const {type, operator} = attr;
  switch (type) {
    case 'containsText':
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return {
        type: 'containsText',
        operator: type,
      };

    default:
      return {type, operator};
  }
};

class CfRuleXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      dataBar: this.databarXform = new DatabarXform(),
      extLst: this.extLstXform = new ExtLstRefXform(),
      formula: this.formulaXform = new FormulaXform(),
      colorScale: this.colorScaleXform = new ColorScaleXform(),
      iconSet: this.iconSetXform = new IconSetXform(),
    };
  }

  get tag() {
    return 'cfRule';
  }

  static isPrimitive(rule) {
    // is this rule primitive?
    if (rule.type === 'iconSet') {
      if (rule.custom || extIcons[rule.iconSet]) {
        return false;
      }
    }
    return true;
  }


  render(xmlStream, model) {
    switch (model.type) {
      case 'expression':
        this.renderExpression(xmlStream, model);
        break;
      case 'cellIs':
        this.renderCellIs(xmlStream, model);
        break;
      case 'top10':
        this.renderTop10(xmlStream, model);
        break;
      case 'aboveAverage':
        this.renderAboveAverage(xmlStream, model);
        break;
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'colorScale':
        this.renderColorScale(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
      case 'containsText':
        this.renderText(xmlStream, model);
        break;
      case 'timePeriod':
        this.renderTimePeriod(xmlStream, model);
        break;
    }
  }

  renderExpression(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'expression',
      dxfId: model.dxfId,
      priority: model.priority,
    });

    this.formulaXform.render(xmlStream, model.formulae[0]);

    xmlStream.closeNode();
  }

  renderCellIs(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'cellIs',
      dxfId: model.dxfId,
      priority: model.priority,
      operator: model.operator,
    });

    model.formulae.forEach(formula => {
      this.formulaXform.render(xmlStream, formula);
    });

    xmlStream.closeNode();
  }

  renderTop10(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      type: 'top10',
      dxfId: model.dxfId,
      priority: model.priority,
      percent: model.percent ? '1' : undefined,
      bottom: model.bottom ? '1' : undefined,
      rank: model.rank || 10,
    });
  }

  renderAboveAverage(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      type: 'aboveAverage',
      dxfId: model.dxfId,
      priority: model.priority,
      aboveAverage: (model.aboveAverage === false) ? '0' : undefined,
    });
  }

  renderDataBar(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      priority: model.priority,
    });

    this.databarXform.render(xmlStream, model);

    model.x14Id = `{${uuid.v4()}}`;
    this.extLstXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderColorScale(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'colorScale',
      priority: model.priority,
    });

    this.colorScaleXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream, model) {
    // iconset is all primitive or all extLst
    if (!CfRuleXform.isPrimitive(model)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
    });

    this.iconSetXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderText(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: model.operator,
      dxfId: model.dxfId,
      priority: model.priority,
      operator: model.operator === 'containsText' ? model.operator : undefined,
    });

    const formula = getTextFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  renderTimePeriod(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'timePeriod',
      dxfId: model.dxfId,
      priority: model.priority,
      timePeriod: model.timePeriod,
    });

    const formula = getTimePeriodFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {
          ...opType(node.attributes),
          dxfId: BaseXform.toIntValue(node.attributes.dxfId),
          priority: BaseXform.toIntValue(node.attributes.priority),
          timePeriod: node.attributes.timePeriod,
          percent: BaseXform.toBoolValue(node.attributes.percent),
          bottom: BaseXform.toBoolValue(node.attributes.bottom),
          rank: BaseXform.toIntValue(node.attributes.rank),
          aboveAverage: BaseXform.toBoolValue(node.attributes.aboveAverage),
        };
        return true;

      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
    }
    return false;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        switch(name) {
          case 'dataBar':
          case 'extLst':
          case 'colorScale':
          case 'iconSet':
            // merge parser model with ours
            Object.assign(this.model, this.parser.model);
            break;

          case 'formula':
            // except - formula is a string and appends to formulae
            this.model.formulae = this.model.formulae || [];
            this.model.formulae.push(this.parser.model);
            break;
        }

        this.parser = null;
      }
      return true;
    }
    return name !== this.tag;
  }
}

module.exports = CfRuleXform;
