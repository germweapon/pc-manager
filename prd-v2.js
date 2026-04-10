const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, LevelFormat } = require('docx');
const fs = require('fs');

const ACCENT = '6C5CE7';
const GRAY = '666666';
const WHITE = 'FFFFFF';
const LIGHT_BG = 'F0F0F5';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

function hCell(t, w) {
  return new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cm,
    children: [new Paragraph({ children: [new TextRun({ text: t, size: 20, bold: true, font: 'Arial', color: WHITE })] })] });
}
function dCell(t, w, shade) {
  return new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: shade ? { fill: LIGHT_BG, type: ShadingType.CLEAR } : undefined, margins: cm,
    children: [new Paragraph({ children: [new TextRun({ text: String(t), size: 20, font: 'Arial' })] })] });
}
function tbl(hdrs, rows, cw) {
  const tw = cw.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: cw,
    rows: [
      new TableRow({ children: hdrs.map((h, i) => hCell(h, cw[i])) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, ci) => dCell(c, cw[ci], ri % 2 === 1)) }))
    ]
  });
}
function p(text, opts) {
  return new Paragraph({ spacing: { after: 120 }, ...opts,
    children: [new TextRun({ size: 22, font: 'Arial', ...(opts?.run || {}), text })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ size: 22, font: 'Arial', text })] });
}

const sections = [];

// Cover
sections.push({
  properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
  children: [
    new Paragraph({ spacing: { before: 4000 }, children: [] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: 'iSens PC Bang Manager', size: 52, bold: true, font: 'Arial', color: ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
      children: [new TextRun({ text: 'Product Requirements Document (PRD)', size: 28, font: 'Arial', color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
      children: [new TextRun({ text: 'Version 1.0  |  2026.03.19', size: 22, font: 'Arial', color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'iSens League', size: 22, font: 'Arial', color: GRAY })] }),
  ]
});

// Main content
const children = [];

// 1. Overview
children.push(
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('1. \uC81C\uD488 \uAC1C\uC694')] }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.1 \uC81C\uD488 \uBE44\uC804')] }),
  p('iSens PC Bang Manager\uB294 PC\uBC29(\uC778\uD130\uB137 \uCE74\uD398) \uC6B4\uC601\uC790\uB97C \uC704\uD55C \uD1B5\uD569 \uAD00\uB9AC \uC194\uB8E8\uC158\uC73C\uB85C, \uC2E4\uC2DC\uAC04 PC \uC0C1\uD0DC \uBAA8\uB2C8\uD130\uB9C1, \uC88C\uC11D \uBC30\uCE58 \uAD00\uB9AC, \uAC00\uB3D9\uB960 \uBD84\uC11D, \uC138\uB9AC\uBA38\uB2C8(\uC774\uBCA4\uD2B8 \uD6A8\uACFC) \uAD00\uB9AC \uB4F1 PC\uBC29 \uC6B4\uC601\uC5D0 \uD544\uC694\uD55C \uBAA8\uB4E0 \uAE30\uB2A5\uC744 \uB2E8\uC77C \uC6F9 \uC778\uD130\uD398\uC774\uC2A4\uC5D0\uC11C \uC81C\uACF5\uD569\uB2C8\uB2E4.'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.2 \uC81C\uD488 \uBAA9\uD45C')] }),
  bullet('PC\uBC29 \uB0B4 \uBAA8\uB4E0 PC\uC758 \uC2E4\uC2DC\uAC04 \uC0C1\uD0DC(ON/OFF, \uC628\uB3C4, \uACBD\uACE0)\uB97C \uC2DC\uAC01\uC801\uC73C\uB85C \uBAA8\uB2C8\uD130\uB9C1'),
  bullet('\uB3C4\uBA74 \uAE30\uBC18 \uC9C1\uAD00\uC801 \uC88C\uC11D \uBC30\uCE58 \uBC0F \uAD6C\uC5ED \uAD00\uB9AC'),
  bullet('\uC2DC\uAC04\uB300\uBCC4 \uAC00\uB3D9\uB960 \uBD84\uC11D\uC744 \uD1B5\uD55C \uC6B4\uC601 \uD6A8\uC728\uD654'),
  bullet('\uC138\uB9AC\uBA38\uB2C8(\uAC8C\uC784 \uC774\uBCA4\uD2B8 \uD6A8\uACFC) \uC124\uC815 \uBC0F \uD1B5\uACC4 \uAD00\uB9AC'),
  bullet('\uB2E4\uC911 \uB9E4\uC7A5 \uD1B5\uD569 \uAD00\uB9AC \uC9C0\uC6D0'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('1.3 \uAE30\uC220 \uC2A4\uD0DD')] }),
  tbl(['\uD56D\uBAA9', '\uC0C1\uC138'], [
    ['\uD504\uB860\uD2B8\uC5D4\uB4DC', 'HTML5, CSS3, Vanilla JavaScript'],
    ['\uB514\uC790\uC778 \uC2DC\uC2A4\uD15C', '\uB2E4\uD06C \uD14C\uB9C8, CSS Custom Properties, Pretendard \uD3F0\uD2B8'],
    ['\uC544\uC774\uCF58', 'Remix Icon (CDN)'],
    ['\uCC28\uD2B8', 'Canvas API \uAE30\uBC18 \uC790\uCCB4 \uAD6C\uD604'],
    ['\uBC30\uD3EC', '\uB2E8\uC77C HTML \uD30C\uC77C (Single File Application)'],
  ], [3000, 6360]),

  new Paragraph({ children: [new PageBreak()] }),

  // 2. Auth
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('2. \uC0AC\uC6A9\uC790 \uC778\uC99D \uBC0F \uAD00\uB9AC')] }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.1 \uB85C\uADF8\uC778')] }),
  p('\uC0AC\uC6A9\uC790\uB294 \uC544\uC774\uB514\uC640 \uBE44\uBC00\uBC88\uD638\uB85C \uC2DC\uC2A4\uD15C\uC5D0 \uC811\uADFC\uD569\uB2C8\uB2E4.'),
  bullet('\uC544\uC774\uB514/\uBE44\uBC00\uBC88\uD638 \uC778\uC99D'),
  bullet('\uBE44\uBC00\uBC88\uD638 \uD45C\uC2DC/\uC228\uAE40 \uD1A0\uAE00'),
  bullet('\uB85C\uADF8\uC778 \uC0C1\uD0DC \uC720\uC9C0 (\uC138\uC158 \uC2A4\uD1A0\uB9AC\uC9C0)'),
  bullet('\uD68C\uC6D0\uAC00\uC785 \uD398\uC774\uC9C0 \uB9C1\uD06C'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.2 \uD68C\uC6D0\uAC00\uC785')] }),
  p('\uC2E0\uADDC \uC0AC\uC6A9\uC790\uB294 \uAC1C\uC778\uC815\uBCF4\uB97C \uC785\uB825\uD558\uC5EC \uAC00\uC785\uC744 \uC2E0\uCCAD\uD569\uB2C8\uB2E4. \uAD00\uB9AC\uC790 \uC2B9\uC778 \uD6C4 \uB85C\uADF8\uC778\uC774 \uAC00\uB2A5\uD569\uB2C8\uB2E4.'),
  tbl(['\uD544\uB4DC', '\uD544\uC218 \uC5EC\uBD80', '\uC124\uBA85'], [
    ['\uD68C\uC6D0 \uAD6C\uBD84', '\uD544\uC218', '\uB4DC\uB86D\uB2E4\uC6B4 \uC120\uD0DD'],
    ['\uC544\uC774\uB514', '\uD544\uC218', '\uACE0\uC720 \uC2DD\uBCC4\uC790'],
    ['\uBE44\uBC00\uBC88\uD638', '\uD544\uC218', '\uAC15\uB3C4 \uD45C\uC2DC (\uC57D/\uBCF4\uD1B5/\uAC15)'],
    ['\uC774\uB984', '\uD544\uC218', '\uC2E4\uBA85'],
    ['\uD734\uB300\uD3F0 \uBC88\uD638', '\uD544\uC218', '\uC5F0\uB77D\uCC98'],
    ['E-mail', '\uD544\uC218', '\uC774\uBA54\uC77C \uC8FC\uC18C'],
    ['\uC18C\uC18D / \uC9C1\uAE09', '\uC120\uD0DD', '\uC18C\uC18D \uBD80\uC11C, \uC9C1\uAE09'],
  ], [2400, 1500, 5460]),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('2.3 \uAC1C\uC778\uC815\uBCF4\uAD00\uB9AC')] }),
  bullet('\uACC4\uC815 \uC815\uBCF4: \uC544\uC774\uB514(\uC77D\uAE30\uC804\uC6A9), \uBE44\uBC00\uBC88\uD638 \uBCC0\uACBD'),
  bullet('\uAC1C\uC778 \uC815\uBCF4: \uC774\uB984, \uD734\uB300\uD3F0 \uBC88\uD638, E-mail \uC218\uC815'),
  bullet('\uC18C\uC18D \uC815\uBCF4: \uC18C\uC18D, \uC9C1\uAE09, \uD68C\uC6D0 \uAD6C\uBD84(\uC77D\uAE30\uC804\uC6A9), \uC0C1\uD0DC(\uC77D\uAE30\uC804\uC6A9), \uBA54\uBAA8'),

  new Paragraph({ children: [new PageBreak()] }),

  // 3. PC Management
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('3. PC \uAD00\uB9AC (\uBA54\uC778 \uB300\uC2DC\uBCF4\uB4DC)')] }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.1 \uD654\uBA74 \uAD6C\uC131')] }),
  p('PC \uAD00\uB9AC\uB294 \uC2DC\uC2A4\uD15C\uC758 \uBA54\uC778 \uD654\uBA74\uC73C\uB85C, \uC120\uD0DD\uB41C \uB9E4\uC7A5\uC758 \uBAA8\uB4E0 PC \uC0C1\uD0DC\uB97C \uC2E4\uC2DC\uAC04\uC73C\uB85C \uBAA8\uB2C8\uD130\uB9C1\uD569\uB2C8\uB2E4.'),
  bullet('\uC0C1\uB2E8\uBC14: \uB9E4\uC7A5 \uC120\uD0DD \uB4DC\uB86D\uB2E4\uC6B4, PC \uAD00\uB9AC/\uAC00\uB3D9\uB960 \uD0ED, \uC0C1\uD0DC \uC694\uC57D, \uBDF0 \uD1A0\uAE00'),
  bullet('\uC88C\uCE21 \uC0AC\uC774\uB4DC\uBC14: \uBA54\uB274 \uB124\uBE44\uAC8C\uC774\uC158, \uD504\uB85C\uD544 \uC601\uC5ED'),
  bullet('\uBA54\uC778 \uC601\uC5ED: 3\uAC00\uC9C0 \uBDF0(\uB3C4\uBA74/\uAD6C\uC5ED/\uBAA9\uB85D)\uB85C PC \uC88C\uC11D \uD45C\uC2DC'),
  bullet('\uC6B0\uCE21 \uD328\uB110: \uC120\uD0DD\uB41C PC\uC758 \uC0C1\uC138 \uC815\uBCF4'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.2 \uC0C1\uD0DC \uD45C\uC2DC \uCCB4\uACC4')] }),
  tbl(['\uC0C1\uD0DC', '\uC870\uAC74', '\uC0C9\uC0C1', '\uC124\uBA85'], [
    ['\uAC00\uB3D9(ON)', 'PC \uC804\uC6D0 ON', '#00CEC9 (Cyan)', '\uC815\uC0C1 \uAC00\uB3D9 \uC911\uC778 PC'],
    ['\uC8FC\uC758(Warning)', 'CPU 70-85\u00B0C / GPU 65-80\u00B0C', '#FDCB6E (Yellow)', '\uC628\uB3C4 \uC8FC\uC758 \uD544\uC694'],
    ['\uACBD\uACE0(Error)', 'CPU 85\u00B0C+ / GPU 80\u00B0C+', '#FF6B6B (Red)', '\uC989\uC2DC \uC870\uCE58 \uD544\uC694'],
    ['\uBBF8\uC0AC\uC6A9(OFF)', 'PC \uC804\uC6D0 OFF', '#6B7185 (Gray)', '\uBE44\uAC00\uB3D9 PC'],
  ], [1800, 2600, 2400, 2560]),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.3 \uB3C4\uBA74 \uBDF0 (Floor Plan View)')] }),
  bullet('\uAD6C\uC5ED\uBCC4 \uC0C9\uC0C1 \uCF54\uB529\uB41C \uC88C\uC11D \uCE74\uB4DC'),
  bullet('PC \uBC88\uD638, \uC0C1\uD0DC \uC778\uB514\uCF00\uC774\uD130, CPU \uC628\uB3C4 \uD45C\uC2DC'),
  bullet('\uC90C \uC778/\uC544\uC6C3 \uCEE8\uD2B8\uB864 (50%~200%)'),
  bullet('\uC88C\uC11D \uD638\uBC84 \uC2DC \uC0C1\uC138 \uD234\uD301 (IP, CPU, GPU, RAM)'),
  bullet('\uC88C\uC11D \uD074\uB9AD \uC2DC \uC6B0\uCE21 \uC0C1\uC138 \uD328\uB110 \uC624\uD508'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.4 \uAD6C\uC5ED \uBDF0 (Zone View)')] }),
  bullet('\uAD6C\uC5ED \uD5E4\uB354\uC5D0 \uAD6C\uC5ED\uBA85\uACFC \uC88C\uC11D \uC218 \uD45C\uC2DC'),
  bullet('\uAC01 \uAD6C\uC5ED \uB0B4 \uC88C\uC11D\uC744 \uAC00\uB85C \uC815\uB82C\uB85C \uBC30\uCE58'),
  bullet('\uD3B8\uC9D1 \uBAA8\uB4DC\uC5D0\uC11C \uAD6C\uC5ED \uAC04 \uC88C\uC11D \uB4DC\uB798\uADF8&\uB4DC\uB86D \uC774\uB3D9'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.5 \uBAA9\uB85D \uBDF0 (List View)')] }),
  bullet('PC\uBC88\uD638, \uAD6C\uC5ED \uCEEC\uB7FC \uC88C\uCE21 \uACE0\uC815 (\uAC00\uB85C \uC2A4\uD06C\uB864 \uC2DC)'),
  bullet('\uD398\uC774\uC9C0\uB124\uC774\uC158: \uD398\uC774\uC9C0\uB2F9 20~200\uAC1C (20\uAC1C \uB2E8\uC704 \uC120\uD0DD), \uC0C1\uD558\uB2E8 \uC6B0\uCE21 \uC815\uB82C'),
  bullet('\uCEEC\uB7FC: \uC0C1\uD0DC, IP, CPU/GPU \uBAA8\uB378, \uC628\uB3C4, \uBC84\uC804, \uC5C5\uB370\uC774\uD2B8'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.6 PC \uC0C1\uC138 \uC815\uBCF4 \uD328\uB110')] }),
  bullet('\uAE30\uBCF8 \uC815\uBCF4: PC\uBC88\uD638, \uAD6C\uC5ED, IP, \uBC84\uC804, RAM, \uC7A5\uCE58 \uC218'),
  bullet('CPU/GPU \uC139\uC158: \uBAA8\uB378\uBA85, \uC628\uB3C4 \uAC8C\uC774\uC9C0, \uD604\uC7AC/\uCD5C\uACE0/\uACBD\uACE0 \uD69F\uC218'),
  bullet('\uC5F0\uACB0 \uC7A5\uCE58: Mouse, Keyboard, Headset \uB4F1 \uD0DC\uADF8 \uD45C\uC2DC'),
  bullet('24\uC2DC\uAC04 \uAC00\uB3D9\uB960: \uC2DC\uAC04\uB300\uBCC4 \uBC14 \uCC28\uD2B8, \uD3C9\uADE0 \uAC00\uB3D9\uB960(%) \uD45C\uC2DC'),
  bullet('\uD558\uB2E8 \uACE0\uC815 \uD234\uBC14: \uC774\uC804/\uC218\uC815/\uB2E4\uC74C \uBC84\uD2BC'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('3.7 \uD3B8\uC9D1 \uBAA8\uB4DC')] }),
  bullet('\uAD6C\uC5ED \uCE74\uB4DC \uB4DC\uB798\uADF8&\uB4DC\uB86D\uC73C\uB85C \uC21C\uC11C \uBCC0\uACBD'),
  bullet('\uC88C\uC11D \uB4DC\uB798\uADF8\uB85C \uAD6C\uC5ED \uAC04 \uC774\uB3D9'),
  bullet('Shift+\uD074\uB9AD\uC73C\uB85C \uB2E4\uC911 \uC120\uD0DD'),
  bullet('\uC6B0\uD074\uB9AD \uCEE8\uD14D\uC2A4\uD2B8 \uBA54\uB274 (\uC218\uC815/\uC0AD\uC81C)'),
  bullet('\uC88C\uC11D \uCD94\uAC00/\uC0AD\uC81C'),
  bullet('PC \uAD00\uB9AC, \uAC00\uB3D9\uB960 \uD0ED \uBAA8\uB450\uC5D0\uC11C \uD3B8\uC9D1 \uAC00\uB2A5'),

  new Paragraph({ children: [new PageBreak()] }),

  // 4. Utilization
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('4. \uAC00\uB3D9\uB960 \uBD84\uC11D')] }),
  p('\uB3D9\uC77C\uD55C 3\uAC00\uC9C0 \uBDF0\uC5D0\uC11C \uC2DC\uAC04\uB300\uBCC4 PC \uAC00\uB3D9\uB960\uC744 \uC2DC\uAC01\uC801\uC73C\uB85C \uBD84\uC11D\uD569\uB2C8\uB2E4. \uB2E8\uC77C \uD37C\uD50C \uCEEC\uB7EC\uC758 \uBA85\uB3C4/\uCC44\uB3C4 \uADF8\uB77C\uB370\uC774\uC158\uC73C\uB85C 6\uB2E8\uACC4 \uD45C\uD604\uD569\uB2C8\uB2E4.'),
  tbl(['\uAC00\uB3D9\uB960 \uBC94\uC704', '\uC2DC\uAC01\uC801 \uD2B9\uC131'], [
    ['0% (\uBBF8\uC0AC\uC6A9)', '\uAC70\uC758 \uD22C\uBA85\uD55C \uD68C\uC0C9'],
    ['1~20%', '\uB9E4\uC6B0 \uC5F0\uD55C \uD37C\uD50C'],
    ['21~40%', '\uC5F0\uD55C \uD37C\uD50C'],
    ['41~60%', '\uC911\uAC04 \uD37C\uD50C'],
    ['61~80%', '\uC9C4\uD55C \uD37C\uD50C'],
    ['81~100%', '\uAC00\uC7A5 \uC9C4\uD55C \uD37C\uD50C'],
  ], [4680, 4680]),

  bullet('\uD604\uC7AC \uC2DC\uAC04\uB300 \uD45C\uC2DC (\uC608: 13~14\uC2DC)'),
  bullet('\uD3C9\uADE0 \uAC00\uB3D9\uB960 (%) \uD45C\uC2DC'),
  bullet('\uB370\uC774\uD130 \uAE30\uC900 \uC2DC\uAC04 \uD45C\uC2DC'),
  bullet('\uBAA9\uB85D \uBDF0: 24\uC2DC\uAC04 \uC2DC\uAC04\uB300\uBCC4 \uCEEC\uB7FC \uC0C9\uC0C1 \uCF54\uB529'),

  new Paragraph({ children: [new PageBreak()] }),

  // 5. Store Management
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('5. \uB9E4\uC7A5 \uAD00\uB9AC')] }),
  p('\uB4F1\uB85D\uB41C \uBAA8\uB4E0 \uB9E4\uC7A5\uC744 \uD14C\uC774\uBE14 \uD615\uD0DC\uB85C \uC870\uD68C\uD558\uACE0 \uAD00\uB9AC\uD569\uB2C8\uB2E4.'),
  bullet('\uB9E4\uC7A5\uCF54\uB4DC, \uB9E4\uC7A5\uBA85 \uCEEC\uB7FC \uC88C\uCE21 \uACE0\uC815'),
  bullet('\uB9E4\uC7A5 \uB4F1\uB85D: \uB9E4\uC7A5\uBA85, \uB300\uD45C\uC790\uBA85, \uC0AC\uC5C5\uC790\uB4F1\uB85D\uBC88\uD638, \uAC00\uB9F9\uC810\uC8FC, \uBCF8\uC0AC SV, CPU/GPU \uC54C\uB78C\uAE30\uC900\uC628\uB3C4, \uB3C4\uBA74 \uC774\uBBF8\uC9C0 \uC5C5\uB85C\uB4DC'),
  bullet('\uB9E4\uC7A5 \uD589 \uD074\uB9AD \uC2DC \uC218\uC815 \uBAA8\uB2EC \uD45C\uC2DC'),
  bullet('\uB9E4\uC7A5 \uC0AD\uC81C \uAE30\uB2A5'),

  new Paragraph({ children: [new PageBreak()] }),

  // 6. PC Registration
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('6. PC \uB4F1\uB85D/\uC218\uC815')] }),
  p('\uC120\uD0DD\uB41C \uB9E4\uC7A5\uC758 \uBAA8\uB4E0 PC\uB97C \uAD00\uB9AC\uD569\uB2C8\uB2E4.'),
  bullet('\uCCB4\uD06C\uBC15\uC2A4/PC\uC0C1\uD0DC/PC\uBC88\uD638/IP \uCEEC\uB7FC \uC88C\uCE21 \uACE0\uC815'),
  bullet('\uC804\uCCB4 PC / \uAC00\uB3D9 PC / \uBBF8\uC0AC\uC6A9 PC \uD544\uD130'),
  bullet('\uC804\uCCB4 \uC120\uD0DD/\uD574\uC81C \uCCB4\uD06C\uBC15\uC2A4'),
  bullet('\uC120\uD0DDPC \uCCB4\uD06C / \uC120\uD0DDPC \uC0AD\uC81C / PC\uB4F1\uB85D / IP\uC218\uC815 \uBC84\uD2BC'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('6.1 PC \uC138\uBD80\uB0B4\uC6A9 \uC218\uC815 \uBAA8\uB2EC')] }),
  tbl(['\uD544\uB4DC', '\uC218\uC815 \uAC00\uB2A5', '\uC124\uBA85'], [
    ['PC\uBC88\uD638 / \uAD6C\uC5ED / IP', 'O', '\uAE30\uBCF8 \uC2DD\uBCC4 \uC815\uBCF4'],
    ['iSensManager Ver. / RAM', 'O', '\uC18C\uD504\uD2B8\uC6E8\uC5B4 \uBC0F \uD558\uB4DC\uC6E8\uC5B4'],
    ['\uC5F0\uACB0 \uC7A5\uCE58 / CPU/GPU \uC0AC\uC591', 'O', '\uC7A5\uCE58 \uBC0F \uBAA8\uB378\uBA85'],
    ['CPU/GPU \uC628\uB3C4 / \uCD5C\uACE0 / \uACBD\uACE0', 'X (\uC77D\uAE30\uC804\uC6A9)', '\uC2E4\uC2DC\uAC04 \uB370\uC774\uD130'],
  ], [3200, 1400, 4760]),

  new Paragraph({ children: [new PageBreak()] }),

  // 7. Ceremony
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('7. \uC138\uB9AC\uBA38\uB2C8 \uAD00\uB9AC')] }),
  p('\uAC8C\uC784 \uC774\uBCA4\uD2B8(\uC2B9\uB9AC, \uD0AC, MVP \uB4F1) \uBC1C\uC0DD \uC2DC \uD300\uB8F8\uC758 Shelly IoT \uC7A5\uCE58\uB97C \uD1B5\uD574 \uC870\uBA85/\uC74C\uD5A5 \uD6A8\uACFC\uB97C \uC2E4\uD589\uD558\uB294 \uAE30\uB2A5\uC785\uB2C8\uB2E4.'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('7.1 \uD1B5\uACC4 \uD0ED')] }),
  bullet('\uAE30\uAC04 \uB0B4 \uC138\uB9AC\uBA38\uB2C8 \uC2E4\uD589\uC218/\uC911\uB2E8\uC218 \uCE74\uB4DC'),
  bullet('\uC2DC\uAC04/\uC77C/\uC8FC/\uC6D4 \uD544\uD130 \uD1A0\uAE00 \uBC0F \uB0A0\uC9DC \uBC94\uC704 \uC120\uD0DD'),
  bullet('TOP 5 \uC138\uB9AC\uBA38\uB2C8/\uC74C\uC6D0 \uB7AD\uD0B9 \uD14C\uC774\uBE14'),
  bullet('Canvas \uAE30\uBC18 \uC2E4\uD589/\uC911\uB2E8 \uCD94\uC774 \uB77C\uC778 \uCC28\uD2B8'),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('7.2 \uC124\uC815 \uD0ED')] }),
  p('\uC88C\uCE21 \uC0AC\uC774\uB4DC\uBC14 + \uC6B0\uCE21 \uCF58\uD150\uCE20 \uB808\uC774\uC544\uC6C3\uC73C\uB85C \uD300\uB8F8\uBCC4 \uC138\uB9AC\uBA38\uB2C8\uB97C \uC124\uC815\uD569\uB2C8\uB2E4.'),
  bullet('\uD300\uB8F8 \uBAA9\uB85D: \uCD94\uAC00/\uC0AD\uC81C \uC9C0\uC6D0'),
  bullet('\uC74C\uC6D0 \uAD00\uB9AC: \uCD94\uAC00/\uC218\uC815/\uC0AD\uC81C \uC9C0\uC6D0'),
  bullet('\uAE30\uBCF8 \uC124\uC815: \uB9E4\uC7A5\uBA85, \uD300\uB8F8\uBA85, ShellyIP, DeviceId, \uC5F0\uACB0 \uC0C1\uD0DC'),
  bullet('\uC138\uB808\uBAA8\uB2C8 \uC124\uC815: 6\uAC1C \uC2AC\uB86F, \uAC01\uAC01 \uC138\uB9AC\uBA38\uB2C8\uBA85/\uC74C\uC6D0 \uC120\uD0DD/\uBB3C\uB9AC \uBC84\uD2BC \uD1A0\uAE00'),
  bullet('\uC0AC\uC6B4\uB4DC \uC124\uC815: \uBCFC\uB968 \uAC15\uC81C \uC870\uC815 \uD1A0\uAE00, \uBCFC\uB968 \uC2AC\uB77C\uC774\uB354 (0~100)'),

  new Paragraph({ children: [new PageBreak()] }),

  // 8. Zones
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('8. \uAD6C\uC5ED \uAD00\uB9AC')] }),
  tbl(['\uAD6C\uC5ED\uBA85', '\uC0C9\uC0C1 \uCF54\uB4DC', '\uC6A9\uB3C4'], [
    ['FPS\uC874', '#6C5CE7 (Purple)', 'FPS \uAC8C\uC784 \uC804\uC6A9'],
    ['LOL\uC874', '#E17055 (Orange)', 'LOL \uAC8C\uC784 \uC804\uC6A9'],
    ['VIP\uC874', '#FDCB6E (Yellow)', 'VIP \uACE0\uAC1D \uC804\uC6A9'],
    ['\uD300\uB8F8', '#FF6348 (Red)', '\uB2E8\uCCB4 \uC774\uC6A9 \uD300\uB8F8'],
    ['\uBA40\uD2F0\uC874', '#636E72 (Gray)', '\uC77C\uBC18 \uBA40\uD2F0 \uAC8C\uC784'],
    ['FC ONLINE\uC874', '#00B894 (Green)', 'FC ONLINE \uC804\uC6A9'],
    ['\uCEE4\uD50C\uC874', '#FD79A8 (Pink)', '\uCEE4\uD50C \uC804\uC6A9'],
    ['\uD504\uB80C\uC988\uC874', '#00CEC9 (Cyan)', '\uCE5C\uAD6C \uADF8\uB8F9'],
    ['\uD37C\uC2A4\uD2B8\uD074\uB798\uC2A4\uC874', '#F9CA24 (Gold)', '\uD504\uB9AC\uBBF8\uC5C4'],
    ['\uB374\uD0C8\uC874', '#74B9FF (Blue)', '\uCE58\uACFC/\uC758\uB8CC \uC81C\uD734'],
  ], [2400, 2800, 4160]),
  bullet('\uAD6C\uC5ED \uCD94\uAC00/\uC218\uC815/\uC0AD\uC81C \uAE30\uB2A5'),

  new Paragraph({ children: [new PageBreak()] }),

  // 9. Real-time
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('9. \uC2E4\uC2DC\uAC04 \uB370\uC774\uD130 \uC5C5\uB370\uC774\uD2B8')] }),
  p('8\uCD08 \uAC04\uACA9\uC73C\uB85C \uAC00\uB3D9 \uC911\uC778 PC\uC758 \uC628\uB3C4 \uB370\uC774\uD130\uB97C \uAC31\uC2E0\uD569\uB2C8\uB2E4.'),
  bullet('CPU \uC628\uB3C4: 30~95\u00B0C \uBC94\uC704, \u00B12~5\u00B0C \uB79C\uB364 \uBCC0\uB3D9'),
  bullet('GPU \uC628\uB3C4: 30~90\u00B0C \uBC94\uC704, \u00B12~5\u00B0C \uB79C\uB364 \uBCC0\uB3D9'),
  bullet('\uCD5C\uACE0 \uC628\uB3C4 \uC790\uB3D9 \uAC31\uC2E0, \uACBD\uACE0 \uD69F\uC218 \uB204\uC801'),
  bullet('\uC0C1\uD0DC \uC790\uB3D9 \uBCC0\uACBD (\uC628\uB3C4 \uC784\uACC4\uAC12 \uAE30\uBC18)'),

  new Paragraph({ children: [new PageBreak()] }),

  // 10. Design System
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('10. \uB514\uC790\uC778 \uC2DC\uC2A4\uD15C')] }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('10.1 \uC0C9\uC0C1 \uD314\uB808\uD2B8')] }),
  tbl(['\uBCC0\uC218\uBA85', '\uC0C9\uC0C1 \uCF54\uB4DC', '\uC6A9\uB3C4'], [
    ['--bg-primary', '#0F1117', '\uBA54\uC778 \uBC30\uACBD'],
    ['--bg-secondary', '#1A1D27', '\uC0AC\uC774\uB4DC\uBC14, \uD328\uB110 \uBC30\uACBD'],
    ['--accent', '#6C5CE7', '\uC8FC\uC694 \uAC15\uC870\uC0C9'],
    ['--success', '#00CEC9', '\uC131\uACF5, \uAC00\uB3D9 \uC0C1\uD0DC'],
    ['--warning', '#FDCB6E', '\uC8FC\uC758 \uC0C1\uD0DC'],
    ['--danger', '#FF6B6B', '\uACBD\uACE0, \uC5D0\uB7EC, \uC0AD\uC81C'],
    ['--text-primary', '#E8EAED', '\uC8FC\uC694 \uD14D\uC2A4\uD2B8'],
    ['--text-muted', '#6B7185', '\uBE44\uD65C\uC131 \uD14D\uC2A4\uD2B8'],
  ], [2400, 2000, 4960]),

  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('10.2 \uD0C0\uC774\uD3EC\uADF8\uB798\uD53C \uBC0F \uCEF4\uD3EC\uB10C\uD2B8')] }),
  bullet('\uAE30\uBCF8 \uD3F0\uD2B8: Pretendard (\uD55C\uAE00 \uCD5C\uC801\uD654)'),
  bullet('\uBC84\uD2BC: .btn (\uAE30\uBCF8), .btn-primary, .btn-danger, .btn-success'),
  bullet('\uBAA8\uB2EC: \uC624\uBC84\uB808\uC774 + \uC911\uC559 \uD328\uB110, \uD5E4\uB354/\uBC14\uB514/\uD478\uD130'),
  bullet('\uD1A0\uC2A4\uD2B8: \uD558\uB2E8 \uC911\uC559, 2.5\uCD08 \uC790\uB3D9 \uC0AC\uB77C\uC9D0'),
  bullet('\uD1A0\uAE00 \uC2A4\uC704\uCE58: 44x24px, \uC2AC\uB77C\uC774\uB529 \uB178\uBE0C'),
);

sections.push({
  properties: {
    page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
  },
  headers: {
    default: new Header({ children: [
      new Paragraph({ alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: ACCENT, space: 1 } },
        children: [new TextRun({ text: 'iSens PC Bang Manager - PRD v1.0', size: 16, font: 'Arial', color: GRAY, italics: true })] })
    ]})
  },
  footers: {
    default: new Footer({ children: [
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Page ', size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY })] })
    ]})
  },
  children
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/Users/isens/Claude/pc-manager/iSens_PC_Bang_Manager_PRD_v1.0.docx', buffer);
  console.log('PRD v2 created:', buffer.length, 'bytes');
});
