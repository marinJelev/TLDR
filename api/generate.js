const pptxgen = require('pptxgenjs');

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { allSections, dateStr, dateSlug } = req.body;

    const sectionColors = {
      'AI': '5E5CE6', 'Tech': '0066CC', 'Dev': '34C759',
      'Product': 'FF9500', 'Founders': 'FF3B30', 'Fintech': '30B0C7'
    };

    const pres = new pptxgen();
    pres.layout = 'LAYOUT_WIDE';
    const W = 13.3, H = 7.5;

    // COVER SLIDE
    const cover = pres.addSlide();
    cover.background = { color: '000000' };
    cover.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.7, y: 1.6, w: 1.9, h: 0.45,
      fill: { color: '0066CC' }, rectRadius: 0.08, line: { color: '0066CC' }
    });
    cover.addText('TLDR DIGEST', {
      x: 5.7, y: 1.6, w: 1.9, h: 0.45,
      fontSize: 10, bold: true, color: 'FFFFFF',
      align: 'center', valign: 'middle', margin: 0, fontFace: 'Arial'
    });
    cover.addText('Your Daily\nTech Briefing', {
      x: 1.5, y: 2.2, w: 10.3, h: 2.0,
      fontSize: 60, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Arial'
    });
    cover.addText(dateStr, {
      x: 1.5, y: 4.35, w: 10.3, h: 0.5,
      fontSize: 17, color: '86868B', align: 'center', fontFace: 'Arial'
    });

    const sectionNames = ['AI', 'Tech', 'Dev', 'Product', 'Founders', 'Fintech'];
    const dotSpacing = 1.5;
    const startX = (W - (sectionNames.length - 1) * dotSpacing) / 2;
    sectionNames.forEach((name, i) => {
      const x = startX + i * dotSpacing - 0.3;
      cover.addShape(pres.shapes.OVAL, {
        x: x + 0.09, y: 5.2, w: 0.12, h: 0.12,
        fill: { color: sectionColors[name] }, line: { color: sectionColors[name] }
      });
      cover.addText(name, {
        x: x - 0.1, y: 5.38, w: 0.6, h: 0.22,
        fontSize: 9, color: '86868B', align: 'center', fontFace: 'Arial'
      });
    });
    cover.addText('6 Sections  \u2022  Daily Edition', {
      x: 1.5, y: 6.8, w: 10.3, h: 0.3,
      fontSize: 10, color: '444444', align: 'center', fontFace: 'Arial'
    });

    // SECTION + ARTICLE SLIDES
    allSections.forEach((sectionData, sectionIdx) => {
      const section = sectionData.section;
      const articles = sectionData.articles || [];
      const accentColor = sectionColors[section] || '0066CC';

      const headerSlide = pres.addSlide();
      headerSlide.background = { color: 'F5F5F7' };
      headerSlide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 0.1, h: H, fill: { color: accentColor }, line: { color: accentColor }
      });
      headerSlide.addText(String(sectionIdx + 1).padStart(2, '0'), {
        x: 0.5, y: 1.3, w: 3, h: 1.8, fontSize: 96, bold: true,
        color: 'E0E0E5', fontFace: 'Arial', align: 'left'
      });
      headerSlide.addText(section, {
        x: 0.5, y: 2.9, w: 9, h: 1.3, fontSize: 72, bold: true,
        color: '1D1D1F', fontFace: 'Arial', align: 'left'
      });
      headerSlide.addText('Top stories  \u2022  ' + new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric' }), {
        x: 0.5, y: 4.3, w: 8, h: 0.4, fontSize: 14,
        color: '86868B', fontFace: 'Arial', align: 'left'
      });
      for (let d = 0; d < 3; d++) {
        headerSlide.addShape(pres.shapes.OVAL, {
          x: 10.5 + d * 0.55, y: 3.5, w: 0.28, h: 0.28,
          fill: { color: accentColor }, line: { color: accentColor }
        });
      }

      articles.forEach((article, articleIdx) => {
        const slide = pres.addSlide();
        slide.background = { color: 'FFFFFF' };
        slide.addShape(pres.shapes.RECTANGLE, {
          x: 0, y: 0, w: W, h: 0.07, fill: { color: accentColor }, line: { color: accentColor }
        });
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x: 0.6, y: 0.32, w: 1.0, h: 0.32,
          fill: { color: accentColor }, rectRadius: 0.06, line: { color: accentColor }
        });
        slide.addText(section, {
          x: 0.6, y: 0.32, w: 1.0, h: 0.32,
          fontSize: 9, bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle', margin: 0, fontFace: 'Arial'
        });
        slide.addText(String(articleIdx + 1) + ' of ' + articles.length, {
          x: 1.75, y: 0.36, w: 1.5, h: 0.22,
          fontSize: 10, color: 'AEAEB2', fontFace: 'Arial', align: 'left'
        });
        slide.addText(article.headline || 'Latest Update', {
          x: 0.6, y: 0.82, w: 12.1, h: 1.7,
          fontSize: 34, bold: true, color: '1D1D1F',
          fontFace: 'Arial', align: 'left', valign: 'top'
        });
        slide.addShape(pres.shapes.LINE, {
          x: 0.6, y: 2.6, w: 11.8, h: 0, line: { color: 'E8E8ED', width: 1 }
        });
        const summaryText = (article.summary || '')
          .replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
        slide.addText(summaryText, {
          x: 0.6, y: 2.75, w: 12.1, h: 2.2,
          fontSize: 16, color: '1D1D1F',
          fontFace: 'Arial', align: 'left', valign: 'top'
        });
        const whyText = article.why_it_matters || '';
        if (whyText) {
          slide.addShape(pres.shapes.RECTANGLE, {
            x: 0.6, y: 5.1, w: 12.1, h: 0.9,
            fill: { color: 'F5F5F7' }, line: { color: 'E8E8ED' }
          });
          slide.addShape(pres.shapes.RECTANGLE, {
            x: 0.6, y: 5.1, w: 0.06, h: 0.9,
            fill: { color: accentColor }, line: { color: accentColor }
          });
          slide.addText('WHY IT MATTERS', {
            x: 0.85, y: 5.17, w: 4, h: 0.22,
            fontSize: 8, bold: true, color: accentColor, fontFace: 'Arial'
          });
          slide.addText(whyText, {
            x: 0.85, y: 5.42, w: 11.6, h: 0.48,
            fontSize: 13, color: '1D1D1F', fontFace: 'Arial', align: 'left'
          });
        }
        slide.addShape(pres.shapes.LINE, {
          x: 0, y: 7.18, w: W, h: 0, line: { color: 'E8E8ED', width: 0.75 }
        });
        slide.addText('TLDR ' + section, {
          x: 0.6, y: 7.22, w: 6, h: 0.24,
          fontSize: 9, color: 'AEAEB2', fontFace: 'Arial'
        });
      });
    });

    // CLOSING SLIDE
    const finalSlide = pres.addSlide();
    finalSlide.background = { color: '000000' };
    finalSlide.addText('Stay curious.', {
      x: 1.5, y: 2.4, w: 10.3, h: 1.5,
      fontSize: 68, bold: true, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
    });
    finalSlide.addText('See you tomorrow.', {
      x: 1.5, y: 4.1, w: 10.3, h: 0.6,
      fontSize: 20, color: '86868B', fontFace: 'Arial', align: 'center'
    });
    finalSlide.addText('T L D R', {
      x: 1.5, y: 5.5, w: 10.3, h: 0.5,
      fontSize: 12, color: '444444', fontFace: 'Arial', align: 'center'
    });

    const base64 = await pres.write({ outputType: 'base64' });
    const filename = 'TLDR_Digest_' + (dateSlug || new Date().toISOString().split('T')[0]) + '.pptx';

    res.status(200).json({ base64, filename });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
