(function () {
// Converts between collection objects and the Matrixify-style "Collections" sheet layout.
//
// Layout: one row per product (manual collections) or per condition (smart collections).
// The first ("top") row of each collection carries the collection-level fields.

const SORT_LABELS = {
  MANUAL: 'Manual',
  BEST_SELLING: 'Best Selling',
  ALPHA_ASC: 'Alphabetically: A-Z',
  ALPHA_DESC: 'Alphabetically: Z-A',
  PRICE_ASC: 'Price: Low to High',
  PRICE_DESC: 'Price: High to Low',
  CREATED: 'Created',
  CREATED_DESC: 'Created Desc',
  MOST_RELEVANT: 'Most Relevant',
};
const FIELD_LABELS = {
  TAG: 'Tag',
  TITLE: 'Title',
  TYPE: 'Product Type',
  VENDOR: 'Vendor',
  VARIANT_PRICE: 'Price',
  VARIANT_COMPARE_AT_PRICE: 'Compare at Price',
  VARIANT_WEIGHT: 'Weight',
  VARIANT_INVENTORY: 'Inventory Stock',
  VARIANT_TITLE: "Variant's Title",
  IS_PRICE_REDUCED: 'Price Reduced',
  PRODUCT_CATEGORY_ID_WITH_DESCENDANTS: 'Product Category',
};
const REL_LABELS = {
  EQUALS: 'Equals',
  NOT_EQUALS: 'Is not equal to',
  GREATER_THAN: 'Is greater than',
  LESS_THAN: 'Is less than',
  STARTS_WITH: 'Starts with',
  ENDS_WITH: 'Ends with',
  CONTAINS: 'Contains',
  NOT_CONTAINS: 'Does not contain',
  IS_SET: 'Is set',
  IS_NOT_SET: 'Is not set',
};

const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const reverse = (labels, extra = {}) => {
  const m = {};
  for (const [k, v] of Object.entries(labels)) { m[norm(v)] = k; m[norm(k)] = k; }
  for (const [k, v] of Object.entries(extra)) m[norm(k)] = v;
  return m;
};
const SORT_REV = reverse(SORT_LABELS, {
  relevance: 'MOST_RELEVANT',
  alphabetically: 'ALPHA_ASC',
  'alphabetically in descending order': 'ALPHA_DESC',
  'highest price': 'PRICE_DESC',
  'lowest price': 'PRICE_ASC',
});
const FIELD_REV = reverse(FIELD_LABELS, {
  'product tag': 'TAG', 'product title': 'TITLE', 'product type': 'TYPE', type: 'TYPE', 'product vendor': 'VENDOR',
  'product price': 'VARIANT_PRICE', 'variant price': 'VARIANT_PRICE', 'compare price': 'VARIANT_COMPARE_AT_PRICE',
  'compare at price': 'VARIANT_COMPARE_AT_PRICE', 'inventory stock': 'VARIANT_INVENTORY', 'variant s title': 'VARIANT_TITLE',
  'variant title': 'VARIANT_TITLE', 'product weight': 'VARIANT_WEIGHT',
});
const REL_REV = reverse(REL_LABELS, {
  'is equal to': 'EQUALS', 'not equals': 'NOT_EQUALS', 'greater than': 'GREATER_THAN', 'less than': 'LESS_THAN',
  'does not contains': 'NOT_CONTAINS', 'is not empty': 'IS_SET', 'is empty': 'IS_NOT_SET',
});

const MANUAL = 'Manual Selection';

// collections: objects as produced by runExport in server.js
function collectionsToRows(collections) {
  const pubNames = [], mfCols = [];
  for (const c of collections) {
    for (const p of c.publications || []) if (!pubNames.includes(p.name)) pubNames.push(p.name);
    for (const m of c.metafields || []) {
      const col = `Metafield: ${m.namespace}.${m.key} [${m.type}]`;
      if (!mfCols.includes(col)) mfCols.push(col);
    }
  }
  const headers = [
    'ID', 'Handle', 'Command', 'Title', 'Body HTML', 'Sort Order', 'Template Suffix', 'Updated At',
    'Image Src', 'Image Alt Text', 'SEO Title', 'SEO Description', 'Products Count', 'Row #', 'Top Row',
    ...pubNames.flatMap((n) => [`Published: ${n}`, `Published At: ${n}`]),
    'Source: Type', 'Inclusion: Type', 'Inclusion: Match',
    'Condition: Command', 'Condition: Field', 'Condition: Relation', 'Condition: Value',
    ...mfCols,
  ];
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  const rows = [headers];

  for (const c of collections) {
    const details = c.ruleSet
      ? c.ruleSet.rules.map((r) => ({ field: FIELD_LABELS[r.column] || r.column, rel: REL_LABELS[r.relation] || r.relation, value: r.condition }))
      : c.productHandles.map((h) => ({ field: MANUAL, value: h }));
    const n = Math.max(1, details.length);
    for (let i = 0; i < n; i++) {
      const row = new Array(headers.length).fill('');
      const set = (h, v) => { if (v !== undefined && v !== null) row[idx[h]] = v; };
      set('ID', c.id);
      set('Handle', c.handle);
      set('Command', 'MERGE');
      set('Title', c.title);
      set('Row #', i + 1);
      if (i === 0) {
        set('Top Row', 1);
        set('Body HTML', c.descriptionHtml);
        set('Sort Order', SORT_LABELS[c.sortOrder] || c.sortOrder);
        set('Template Suffix', c.templateSuffix);
        set('Updated At', c.updatedAt);
        set('Image Src', c.image?.url);
        set('Image Alt Text', c.image?.altText);
        set('SEO Title', c.seo?.title);
        set('SEO Description', c.seo?.description);
        set('Products Count', c.productsCount);
        for (const p of c.publications || []) {
          set(`Published: ${p.name}`, p.published ? 1 : 0);
          set(`Published At: ${p.name}`, p.publishDate);
        }
        for (const m of c.metafields || []) set(`Metafield: ${m.namespace}.${m.key} [${m.type}]`, m.value);
        if (c.ruleSet) set('Inclusion: Match', c.ruleSet.appliedDisjunctively ? 'Any condition' : 'All conditions');
      }
      const d = details[i];
      if (d) {
        set('Source: Type', 'Products');
        set('Inclusion: Type', 'Include');
        set('Condition: Command', 'MERGE');
        set('Condition: Field', d.field);
        set('Condition: Relation', d.rel);
        set('Condition: Value', d.value);
      }
      rows.push(row);
    }
  }
  return rows;
}

// Returns { collections, warnings }
function rowsToCollections(rows) {
  const warnings = [];
  if (!rows.length) throw new Error('The sheet is empty');
  const headers = rows[0].map((h) => String(h).trim().toLowerCase());
  const col = (name) => headers.indexOf(name.toLowerCase());
  if (col('handle') < 0) throw new Error('No "Handle" column found. Is this a collections export?');

  const pubCols = [], mfCols = [];
  headers.forEach((h, i) => {
    let m;
    if ((m = rows[0][i].match(/^\s*Published:\s*(.+?)\s*$/i))) pubCols.push({ name: m[1], i });
    else if ((m = rows[0][i].match(/^\s*Metafield:\s*([^.\s]+)\.(\S+)\s*\[([^\]]+)\]\s*$/i))) mfCols.push({ namespace: m[1], key: m[2], type: m[3], i });
  });

  const ci = Object.fromEntries(
    ['handle', 'command', 'title', 'body html', 'sort order', 'template suffix', 'image src', 'image alt text', 'seo title',
      'seo description', 'inclusion: match', 'condition: field', 'condition: relation', 'condition: value', 'top row']
      .map((n) => [n, col(n)])
  );
  const tagCol = headers.findIndex((h) => /^metafield:\s*title_tag\b/.test(h));
  const descCol = headers.findIndex((h) => /^metafield:\s*description_tag\b/.test(h));
  const cell = (row, i) => (i >= 0 && row[i] != null ? String(row[i]).trim() : '');

  const groups = new Map();
  for (const row of rows.slice(1)) {
    const handle = cell(row, ci.handle);
    if (!handle) continue;
    if (!groups.has(handle)) groups.set(handle, []);
    groups.get(handle).push(row);
  }

  const collections = [];
  for (const [handle, grp] of groups) {
    const top = grp.find((r) => cell(r, ci['top row']) === '1') || grp[0];
    const get = (n) => cell(top, ci[n]);
    const c = {
      handle,
      command: (get('command') || 'MERGE').toUpperCase(),
      title: get('title') || handle,
      descriptionHtml: top[ci['body html']] != null ? String(top[ci['body html']]) : '',
      templateSuffix: get('template suffix') || null,
      productHandles: [],
      ruleSet: null,
      metafields: [],
      published: {},
    };
    const sortRaw = get('sort order');
    if (sortRaw) {
      c.sortOrder = SORT_REV[norm(sortRaw)];
      if (!c.sortOrder) warnings.push(`${handle}: unknown sort order "${sortRaw}", using the store default`);
    }
    const seoTitle = get('seo title') || cell(top, tagCol), seoDesc = get('seo description') || cell(top, descCol);
    if (seoTitle || seoDesc) c.seo = { title: seoTitle || null, description: seoDesc || null };
    if (get('image src')) c.image = { url: get('image src'), altText: get('image alt text') };
    for (const p of pubCols) {
      const v = cell(top, p.i).toLowerCase();
      if (v) c.published[p.name] = ['1', 'true', 'yes'].includes(v);
    }
    for (const m of mfCols) {
      const v = top[m.i];
      if (v !== undefined && v !== '') c.metafields.push({ namespace: m.namespace, key: m.key, type: m.type, value: String(v) });
    }

    const rules = [];
    for (const r of grp) {
      const rowCmd = cell(r, col('condition: command')).toUpperCase();
      if (rowCmd === 'DELETE' || rowCmd === 'IGNORE') continue;
      const field = cell(r, ci['condition: field']);
      const value = cell(r, ci['condition: value']);
      if (!field) continue;
      if (norm(field) === norm(MANUAL)) { if (value) c.productHandles.push(value); continue; }
      const column = FIELD_REV[norm(field)] || (/^[A-Z_]+$/.test(field) ? field : null);
      const relation = REL_REV[norm(cell(r, ci['condition: relation']))] || 'EQUALS';
      if (!column) { warnings.push(`${handle}: unknown condition field "${field}", rule skipped`); continue; }
      rules.push({ column, relation, condition: value });
    }
    if (rules.length) c.ruleSet = { appliedDisjunctively: /any/i.test(get('inclusion: match')), rules };
    collections.push(c);
  }
  return { collections, warnings };
}

const api = { collectionsToRows, rowsToCollections };
if (typeof module !== 'undefined' && module.exports) module.exports = api;
else globalThis.Convert = api;
})();
