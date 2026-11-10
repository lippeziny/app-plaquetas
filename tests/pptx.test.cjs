const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const AdmZip = require('adm-zip');
const { inspectTemplate, buildPptx } = require('../server/lib/pptx.cjs');

const templatePath = path.join(__dirname, '..', 'server', 'assets', 'default-template.pptx');

test('detects both label layouts in the original presentation', () => {
  const template = inspectTemplate(templatePath);
  assert.equal(template.slides.length, 2);
  assert.equal(template.slides[0].capacity, 8);
  assert.equal(template.slides[1].capacity, 8);
});

test('fills only the detected code/name fields and creates enough pages', () => {
  const template = inspectTemplate(templatePath);
  const items = Array.from({ length: 10 }, (_, i) => ({ code: `76100000${i}`, name: `ITEM DE TESTE ${i}` }));
  const output = buildPptx({ templatePath, slideIndex: 0, items });
  assert.equal(output.pages, 2);
  assert.equal(output.capacity, 8);
  const zip = new AdmZip(output.buffer);
  const presentation = zip.getEntry('ppt/presentation.xml').getData().toString('utf8');
  assert.equal((presentation.match(/<p:sldId\b/g) || []).length, 2);
  const slide = zip.getEntry('ppt/slides/slide101.xml').getData().toString('utf8');
  assert.match(slide, /761000000/);
  assert.match(slide, /ITEM DE TESTE 0/);
  assert.match(slide, /Quantidade:/);
  assert.match(slide, /Data de contagem:/);
});
