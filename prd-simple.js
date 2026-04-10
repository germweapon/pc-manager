const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, LevelFormat } = require('docx');
const fs = require('fs');

const A = '6C5CE7';
const G = '888888';
const W = 'FFFFFF';
const b = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const bs = { top: b, bottom: b, left: b, right: b };
const m = { top: 60, bottom: 60, left: 100, right: 100 };

function hc(t, w) {
  return new TableCell({ borders: bs, width: { size: w, type: WidthType.DXA },
    shading: { fill: A, type: ShadingType.CLEAR }, margins: m,
    children: [new Paragraph({ children: [new TextRun({ text: t, size: 18, bold: true, color: W })] })] });
}
function dc(t, w) {
  return new TableCell({ borders: bs, width: { size: w, type: WidthType.DXA }, margins: m,
    children: [new Paragraph({ children: [new TextRun({ text: t, size: 18 })] })] });
}

const doc = new Document({
  numbering: { config: [
    { reference: 'bl', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 260 } } } }] }
  ]},
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
    children: [
      // Title
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: 'iSens PC Bang Manager', size: 36, bold: true, color: A })
      ]}),
      new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: A, space: 1 } },
        children: [new TextRun({ text: 'Product Requirements Document (PRD) v1.0  |  2026.03.19', size: 18, color: G })] }),

      // 1. Overview
      new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 100 },
        children: [new TextRun({ text: '1. \uC81C\uD488 \uAC1C\uC694', size: 26, bold: true })] }),
      new Paragraph({ spacing: { after: 80 }, children: [
        new TextRun({ text: 'PC\uBC29 \uC6B4\uC601\uC790\uB97C \uC704\uD55C \uD1B5\uD569 \uAD00\uB9AC \uC194\uB8E8\uC158\uC73C\uB85C, \uC2E4\uC2DC\uAC04 PC \uBAA8\uB2C8\uD130\uB9C1, \uC88C\uC11D \uBC30\uCE58 \uAD00\uB9AC, \uAC00\uB3D9\uB960 \uBD84\uC11D, \uC138\uB9AC\uBA38\uB2C8 \uAD00\uB9AC \uB4F1\uC744 \uB2E8\uC77C \uC6F9 \uC778\uD130\uD398\uC774\uC2A4\uC5D0\uC11C \uC81C\uACF5\uD569\uB2C8\uB2E4.', size: 20 })
      ]}),

      new Table({ width: { size: 9840, type: WidthType.DXA }, columnWidths: [2400, 7440],
        rows: [
          new TableRow({ children: [hc('\uD56D\uBAA9', 2400), hc('\uC0C1\uC138', 7440)] }),
          new TableRow({ children: [dc('\uAE30\uC220 \uC2A4\uD0DD', 2400), dc('HTML5, CSS3, Vanilla JS, Canvas API', 7440)] }),
          new TableRow({ children: [dc('\uB514\uC790\uC778', 2400), dc('\uB2E4\uD06C \uD14C\uB9C8, Pretendard \uD3F0\uD2B8, Remix Icon', 7440)] }),
          new TableRow({ children: [dc('\uD0C0\uAC9F \uC0AC\uC6A9\uC790', 2400), dc('PC\uBC29 \uC6B4\uC601\uC790, \uBCF8\uC0AC \uAD00\uB9AC\uC790, \uAC00\uB9F9\uC810\uC8FC', 7440)] }),
          new TableRow({ children: [dc('\uBC30\uD3EC \uBC29\uC2DD', 2400), dc('\uB2E8\uC77C HTML \uD30C\uC77C (Single File Application)', 7440)] }),
        ]
      }),

      // 2. Core Features
      new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 100 },
        children: [new TextRun({ text: '2. \uD575\uC2EC \uAE30\uB2A5', size: 26, bold: true })] }),

      new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: '2.1 PC \uAD00\uB9AC (\uBA54\uC778 \uB300\uC2DC\uBCF4\uB4DC)', size: 22, bold: true, color: '333333' })
      ]}),
      ...[
        '\uB3C4\uBA74/\uAD6C\uC5ED/\uBAA9\uB85D 3\uAC00\uC9C0 \uBDF0\uB85C PC \uC88C\uC11D \uC2E4\uC2DC\uAC04 \uBAA8\uB2C8\uD130\uB9C1 (ON/OFF, CPU/GPU \uC628\uB3C4, \uACBD\uACE0)',
        '\uAD6C\uC5ED\uBCC4 \uC0C9\uC0C1 \uCF54\uB529 (10\uAC1C \uAE30\uBCF8 \uAD6C\uC5ED: FPS\uC874, LOL\uC874, VIP\uC874 \uB4F1)',
        '\uC88C\uC11D \uD074\uB9AD \uC2DC \uC6B0\uCE21 \uC0C1\uC138 \uD328\uB110 (PC \uC815\uBCF4, \uC628\uB3C4 \uAC8C\uC774\uC9C0, 24\uC2DC\uAC04 \uAC00\uB3D9\uB960 \uCC28\uD2B8)',
        '\uD3B8\uC9D1 \uBAA8\uB4DC: \uB4DC\uB798\uADF8&\uB4DC\uB86D \uC88C\uC11D \uC774\uB3D9, \uAD6C\uC5ED \uC21C\uC11C \uBCC0\uACBD, \uC88C\uC11D \uCD94\uAC00/\uC0AD\uC81C',
        '8\uCD08 \uAC04\uACA9 \uC2E4\uC2DC\uAC04 \uC628\uB3C4 \uAC31\uC2E0 (CPU 85\u00B0C+/GPU 80\u00B0C+ \uACBD\uACE0)',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: '2.2 \uAC00\uB3D9\uB960 \uBD84\uC11D', size: 22, bold: true, color: '333333' })
      ]}),
      ...[
        '\uB3D9\uC77C\uD55C 3\uAC00\uC9C0 \uBDF0\uC5D0\uC11C \uC2DC\uAC04\uB300\uBCC4 PC \uAC00\uB3D9\uB960 \uC2DC\uAC01\uD654',
        '\uB2E8\uC77C \uD37C\uD50C \uCEEC\uB7EC \uADF8\uB77C\uB370\uC774\uC158 6\uB2E8\uACC4 (0%~100%)',
        '\uBAA9\uB85D \uBDF0: 24\uC2DC\uAC04 \uC2DC\uAC04\uB300\uBCC4 \uCEEC\uB7FC \uC0C9\uC0C1 \uCF54\uB529',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: '2.3 \uB9E4\uC7A5 \uAD00\uB9AC', size: 22, bold: true, color: '333333' })
      ]}),
      ...[
        '\uB2E4\uC911 \uB9E4\uC7A5 \uD1B5\uD569 \uAD00\uB9AC (120\uAC1C \uB9E4\uC7A5 \uC9C0\uC6D0)',
        '\uB9E4\uC7A5 \uB4F1\uB85D/\uC218\uC815/\uC0AD\uC81C, \uB3C4\uBA74 \uC774\uBBF8\uC9C0 \uC5C5\uB85C\uB4DC',
        '\uB9E4\uC7A5\uBCC4 CPU/GPU \uC54C\uB78C\uAE30\uC900\uC628\uB3C4, \uC0AC\uC5C5\uC790\uC815\uBCF4 \uAD00\uB9AC',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: '2.4 PC \uB4F1\uB85D/\uC218\uC815', size: 22, bold: true, color: '333333' })
      ]}),
      ...[
        'PC \uBAA9\uB85D \uC870\uD68C (ON/OFF \uD544\uD130), \uC804\uCCB4 \uC120\uD0DD/\uD574\uC81C',
        'PC \uB4F1\uB85D, IP \uC218\uC815, \uC120\uD0DDPC \uCCB4\uD06C/\uC0AD\uC81C',
        'PC \uC138\uBD80\uB0B4\uC6A9 \uC218\uC815 \uBAA8\uB2EC (\uC628\uB3C4/\uACBD\uACE0 \uD69F\uC218\uB294 \uC77D\uAE30\uC804\uC6A9)',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      new Paragraph({ spacing: { before: 160, after: 60 }, children: [
        new TextRun({ text: '2.5 \uC138\uB9AC\uBA38\uB2C8 \uAD00\uB9AC', size: 22, bold: true, color: '333333' })
      ]}),
      ...[
        '\uD1B5\uACC4 \uD0ED: \uC2E4\uD589\uC218/\uC911\uB2E8\uC218 \uCE74\uB4DC, TOP 5 \uB7AD\uD0B9, \uCD94\uC774 \uB77C\uC778 \uCC28\uD2B8',
        '\uC124\uC815 \uD0ED: \uD300\uB8F8\uBCC4 \uC138\uB808\uBAA8\uB2C8 \uC124\uC815 (6\uAC1C \uC2AC\uB86F, \uC74C\uC6D0 \uC120\uD0DD, Shelly IoT \uC5F0\uB3D9)',
        '\uD300\uB8F8 \uCD94\uAC00/\uC0AD\uC81C, \uC74C\uC6D0 \uB4F1\uB85D/\uC218\uC815/\uC0AD\uC81C, \uBCFC\uB968 \uAC15\uC81C \uC870\uC815',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      // 3. Auth
      new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 100 },
        children: [new TextRun({ text: '3. \uC0AC\uC6A9\uC790 \uAD00\uB9AC', size: 26, bold: true })] }),
      ...[
        '\uB85C\uADF8\uC778: \uC544\uC774\uB514/\uBE44\uBC00\uBC88\uD638, \uBE44\uBC00\uBC88\uD638 \uD45C\uC2DC \uD1A0\uAE00, \uC138\uC158 \uC720\uC9C0',
        '\uD68C\uC6D0\uAC00\uC785: \uD68C\uC6D0\uAD6C\uBD84, \uAC1C\uC778\uC815\uBCF4, \uC18C\uC18D\uC815\uBCF4, \uAD00\uB9AC\uC790 \uC2B9\uC778 \uD6C4 \uB85C\uADF8\uC778',
        '\uAC1C\uC778\uC815\uBCF4\uAD00\uB9AC: \uC774\uB984/\uD578\uB4DC\uD3F0/\uC774\uBA54\uC77C/\uC18C\uC18D/\uC9C1\uAE09 \uC218\uC815, \uBE44\uBC00\uBC88\uD638 \uBCC0\uACBD',
        '\uD504\uB85C\uD544: \uC0AC\uC774\uB4DC\uBC14 \uD558\uB2E8 \uD504\uB85C\uD544 \uC601\uC5ED, \uC124\uC815/\uB85C\uADF8\uC544\uC6C3 \uB4DC\uB86D\uB2E4\uC6B4 \uBA54\uB274',
      ].map(t => new Paragraph({ numbering: { reference: 'bl', level: 0 }, spacing: { after: 40 },
        children: [new TextRun({ text: t, size: 19 })] })),

      // 4. UI/UX
      new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 100 },
        children: [new TextRun({ text: '4. UI/UX \uB514\uC790\uC778 \uC2DC\uC2A4\uD15C', size: 26, bold: true })] }),

      new Table({ width: { size: 9840, type: WidthType.DXA }, columnWidths: [2400, 7440],
        rows: [
          new TableRow({ children: [hc('\uD56D\uBAA9', 2400), hc('\uC0C1\uC138', 7440)] }),
          new TableRow({ children: [dc('\uD14C\uB9C8', 2400), dc('\uB2E4\uD06C \uBAA8\uB4DC (#0F1117 \uBC30\uACBD, #6C5CE7 \uC561\uC13C\uD2B8)', 7440)] }),
          new TableRow({ children: [dc('\uD3F0\uD2B8', 2400), dc('Pretendard (\uD55C\uAE00 \uCD5C\uC801\uD654), -apple-system \uD3F4\uBC31', 7440)] }),
          new TableRow({ children: [dc('\uCEF4\uD3EC\uB10C\uD2B8', 2400), dc('\uBC84\uD2BC, \uBAA8\uB2EC, \uD1A0\uC2A4\uD2B8, \uD1A0\uAE00 \uC2A4\uC704\uCE58, \uB4DC\uB86D\uB2E4\uC6B4, \uD14C\uC774\uBE14', 7440)] }),
          new TableRow({ children: [dc('\uC0C1\uD0DC \uC0C9\uC0C1', 2400), dc('\uC131\uACF5(#00CEC9), \uC8FC\uC758(#FDCB6E), \uACBD\uACE0(#FF6B6B), \uBBF8\uC0AC\uC6A9(#6B7185)', 7440)] }),
          new TableRow({ children: [dc('\uD398\uC774\uC9C0\uB124\uC774\uC158', 2400), dc('\uC6B0\uCE21 \uC815\uB82C, \uD398\uC774\uC9C0\uB2F9 20~200\uAC1C, \uC0C1\uD558\uB2E8 \uB3D9\uC77C \uD45C\uC2DC', 7440)] }),
          new TableRow({ children: [dc('\uACE0\uC815 \uCEEC\uB7FC', 2400), dc('PC\uBC88\uD638/\uAD6C\uC5ED \uCEEC\uB7FC \uC88C\uCE21 \uACE0\uC815 (\uAC00\uB85C \uC2A4\uD06C\uB864 \uC2DC)', 7440)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 200 },
        border: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC', space: 1 } },
        children: [new TextRun({ text: '\u00A9 2026 iSens League. All rights reserved.', size: 16, color: G })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/Users/isens/Claude/pc-manager/iSens_PC_Bang_Manager_PRD_v1.0.docx', buf);
  console.log('Done!', buf.length, 'bytes');
});
