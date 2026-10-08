// Collection Mover: runs entirely in the browser inside the Shopify admin.
// Calls the Admin GraphQL API through App Bridge "direct API access" (no backend, no secrets).
(function () {
const API_VERSION = '2026-07';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function makeClient() {
  async function gql(query, variables = {}) {
    for (let attempt = 0; attempt < 6; attempt++) {
      let r;
      try {
        r = await fetch(`shopify:admin/api/${API_VERSION}/graphql.json`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query, variables }),
        });
      } catch (e) {
        throw new Error(`Could not call the Shopify Admin API (${e.message}). Open this app from the Shopify admin (Apps > Collection Mover) and make sure Direct API access is enabled; see the README.`);
      }
      if (r.status === 429 || r.status >= 500) { await sleep(1000 * (attempt + 1)); continue; }
      if (r.status === 401 || r.status === 403) throw new Error(`Access denied (${r.status}). Reinstall the app and make sure it has the product and publication scopes.`);
      if (!r.ok) throw new Error(`Shopify returned HTTP ${r.status}`);
      const j = await r.json();
      if (j.errors?.length) {
        if (j.errors.some((e) => e.extensions?.code === 'THROTTLED')) { await sleep(1500 * (attempt + 1)); continue; }
        throw new Error(j.errors.map((e) => e.message).join('; '));
      }
      return j.data;
    }
    throw new Error('Too many retries');
  }
  const { shop } = await gql('{ shop { name myshopifyDomain } }');
  return { gql, shop: shop.myshopifyDomain, name: shop.name };
}

// ---------- Export ----------

const LIST_Q = `query($after:String){
  collections(first:20, after:$after){
    pageInfo{hasNextPage endCursor}
    nodes{
      id handle title descriptionHtml sortOrder templateSuffix updatedAt
      productsCount{count}
      resourcePublicationsV2(first:20, onlyPublished:false){nodes{isPublished publishDate publication{name}}}
      seo{title description}
      image{url altText}
      ruleSet{appliedDisjunctively rules{column relation condition}}
      metafields(first:50){nodes{namespace key type value}}
    }
  }
}`;

const PRODUCTS_Q = `query($id:ID!,$after:String){
  collection(id:$id){ products(first:250, after:$after){ pageInfo{hasNextPage endCursor} nodes{handle} } }
}`;

async function runExport(client, log) {
  const out = [];
  let after = null;
  do {
    const d = await client.gql(LIST_Q, { after });
    for (const c of d.collections.nodes) {
      out.push({
        id: c.id.split('/').pop(),
        updatedAt: c.updatedAt,
        productsCount: c.productsCount?.count,
        publications: c.resourcePublicationsV2.nodes.map((n) => ({ name: n.publication.name, published: n.isPublished, publishDate: n.publishDate })),
        handle: c.handle,
        title: c.title,
        descriptionHtml: c.descriptionHtml,
        sortOrder: c.sortOrder,
        templateSuffix: c.templateSuffix,
        seo: c.seo,
        image: c.image,
        ruleSet: c.ruleSet,
        metafields: c.metafields.nodes,
        productHandles: [],
        _id: c.id,
      });
    }
    log(`Fetched ${out.length} collections...`);
    after = d.collections.pageInfo.hasNextPage ? d.collections.pageInfo.endCursor : null;
  } while (after);

  let i = 0;
  for (const c of out) {
    i++;
    if (!c.ruleSet) {
      let a = null;
      do {
        const d = await client.gql(PRODUCTS_Q, { id: c._id, after: a });
        const p = d.collection.products;
        c.productHandles.push(...p.nodes.map((n) => n.handle));
        a = p.pageInfo.hasNextPage ? p.pageInfo.endCursor : null;
      } while (a);
      log(`[${i}/${out.length}] ${c.title}: ${c.productHandles.length} products`);
    } else {
      log(`[${i}/${out.length}] ${c.title}: smart collection (${c.ruleSet.rules.length} rules)`);
    }
    delete c._id;
  }
  return out;
}

async function buildWorkbook(client, collections) {
  const now = new Date();
  const headers = ['Shopify Domain', 'Sheet', 'Exported At', 'Exported'];
  return XlsxLite.writeXlsx([
    { name: 'Collections', rows: Convert.collectionsToRows(collections) },
    { name: 'Export Summary', rows: [headers, [client.shop, 'Collections', now.toISOString(), `${collections.length} of ${collections.length}`]] },
  ]);
}

// ---------- Import ----------

const FIND_Q = `query($q:String!){ collections(first:1, query:$q){ nodes{ id } } }`;
const CREATE_M = `mutation($input:CollectionInput!){ collectionCreate(input:$input){ collection{id} userErrors{field message} } }`;
const UPDATE_M = `mutation($input:CollectionInput!){ collectionUpdate(input:$input){ collection{id} userErrors{field message} } }`;
const ADD_M = `mutation($id:ID!,$ids:[ID!]!){ collectionAddProductsV2(id:$id, productIds:$ids){ userErrors{field message} } }`;
const PUBS_Q = `{ publications(first:25){ nodes{ id name } } }`;
const PUBLISH_M = `mutation($id:ID!,$pub:ID!){ publishablePublish(id:$id, input:[{publicationId:$pub}]){ userErrors{field message} } }`;

async function runImport(client, data, opts, log) {
  const cols = data.collections || [];
  if (!cols.length) throw new Error('The file contains no collections');

  let pubs = [];
  try { pubs = (await client.gql(PUBS_Q)).publications.nodes; } catch (e) { log(`! Could not read sales channels (${e.message}); collections will not be published.`); }
  const pubByName = new Map(pubs.map((p) => [p.name.toLowerCase(), p]));

  const idCache = new Map();
  async function resolveProducts(handles) {
    const need = [...new Set(handles)].filter((h) => !idCache.has(h));
    for (let i = 0; i < need.length; i += 20) {
      const batch = need.slice(i, i + 20);
      const q = `{ ${batch.map((h, k) => `p${k}: productByIdentifier(identifier:{handle:${JSON.stringify(h)}}){id}`).join(' ')} }`;
      const d = await client.gql(q);
      batch.forEach((h, k) => idCache.set(h, d[`p${k}`]?.id || null));
    }
    const ids = [], missing = [];
    for (const h of handles) (idCache.get(h) ? ids : missing).push(idCache.get(h) || h);
    return { ids, missing };
  }

  const stats = { created: 0, updated: 0, skipped: 0, failed: 0, missingProducts: 0 };
  let n = 0;
  for (const c of cols) {
    n++;
    const tag = `[${n}/${cols.length}] ${c.title}`;
    try {
      if (c.command === 'IGNORE' || c.command === 'DELETE') { log(`${tag}: command ${c.command}, skipped`); stats.skipped++; continue; }
      const found = await client.gql(FIND_Q, { q: `handle:${c.handle}` });
      const existing = found.collections.nodes[0]?.id;
      if (existing && opts.existing === 'skip') { log(`${tag}: already exists, skipped`); stats.skipped++; continue; }

      const input = {
        title: c.title,
        handle: c.handle,
        descriptionHtml: c.descriptionHtml || '',
        templateSuffix: c.templateSuffix || null,
        seo: c.seo || undefined,
      };
      if (c.sortOrder) input.sortOrder = c.sortOrder;
      if (c.image?.url) input.image = { src: c.image.url, altText: c.image.altText || '' };

      if (c.ruleSet) {
        const rules = c.ruleSet.rules.filter((r) => r.column !== 'PRODUCT_METAFIELD_DEFINITION');
        const dropped = c.ruleSet.rules.length - rules.length;
        if (dropped) log(`${tag}: ! ${dropped} metafield-definition rule(s) can't be copied between stores`);
        if (!rules.length) { log(`${tag}: no usable rules, skipped`); stats.skipped++; continue; }
        input.ruleSet = { appliedDisjunctively: c.ruleSet.appliedDisjunctively, rules };
      }
      const mf = (c.metafields || []).map(({ namespace, key, type, value }) => ({ namespace, key, type, value }));

      let id = existing;
      const attempt = async (withMeta) => {
        const inp = { ...input, ...(withMeta && mf.length ? { metafields: mf } : {}) };
        if (existing) {
          const r = await client.gql(UPDATE_M, { input: { ...inp, id: existing } });
          return { errs: r.collectionUpdate.userErrors };
        }
        const r = await client.gql(CREATE_M, { input: inp });
        return { errs: r.collectionCreate.userErrors, id: r.collectionCreate.collection?.id };
      };
      let res = await attempt(true);
      if (res.errs.length && mf.length) {
        log(`${tag}: ! retrying without metafields (${res.errs.map((e) => e.message).join('; ')})`);
        res = await attempt(false);
      }
      if (res.errs.length) throw new Error(res.errs.map((e) => e.message).join('; '));
      if (!existing) id = res.id;
      existing ? stats.updated++ : stats.created++;

      let msg = `${tag}: ${existing ? 'updated' : 'created'}`;
      if (!c.ruleSet && c.productHandles?.length) {
        const { ids, missing } = await resolveProducts(c.productHandles);
        for (let i = 0; i < ids.length; i += 250) {
          const r = await client.gql(ADD_M, { id, ids: ids.slice(i, i + 250) });
          const errs = r.collectionAddProductsV2.userErrors;
          if (errs.length) log(`${tag}: ! ${errs.map((e) => e.message).join('; ')}`);
        }
        msg += `, ${ids.length} products added`;
        if (missing.length) { stats.missingProducts += missing.length; msg += `, ${missing.length} not found in this store`; }
      }
      const targets = opts.publishAll ? ['online store'] : Object.keys(c.published || {}).filter((k) => c.published[k]).map((k) => k.toLowerCase());
      const done = [];
      for (const t of targets) {
        const pub = pubByName.get(t);
        if (!pub) { log(`${tag}: ! no sales channel named "${t}" in this store, not published there`); continue; }
        const r = await client.gql(PUBLISH_M, { id, pub: pub.id });
        if (r.publishablePublish.userErrors.length) log(`${tag}: ! publish to ${pub.name}: ${r.publishablePublish.userErrors[0].message}`);
        else done.push(pub.name);
      }
      if (done.length) msg += `, published to ${done.join(', ')}`;
      log(msg);
    } catch (e) {
      stats.failed++;
      log(`${tag}: FAILED - ${e.message}`);
    }
  }
  return stats;
}


// ---------- UI ----------

async function run(sec, job) {
  const btn = sec.querySelector('[data-go]'), logEl = sec.querySelector('[data-log]');
  logEl.style.display = 'block'; logEl.textContent = ''; btn.disabled = true;
  const add = (m) => { logEl.textContent += m + '\n'; logEl.scrollTop = logEl.scrollHeight; };
  try {
    const client = await makeClient();
    add(`Connected to ${client.name} (${client.shop})`);
    await job(client, add);
  } catch (e) { add('ERROR: ' + e.message); }
  btn.disabled = false;
}

const exp = document.getElementById('export');
exp.querySelector('[data-go]').onclick = () =>
  run(exp, async (client, add) => {
    const collections = await runExport(client, add);
    const bytes = await buildWorkbook(client, collections);
    const filename = `Collections_${client.shop.replace('.myshopify.com', '')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = filename;
    a.click();
    add(`\nDone. Exported ${collections.length} collections; ${filename} was downloaded.`);
  });

const imp = document.getElementById('import');
imp.querySelector('[data-go]').onclick = () => {
  const f = document.getElementById('file').files[0];
  if (!f) {
    const l = imp.querySelector('[data-log]');
    l.style.display = 'block'; l.textContent = 'Choose the exported .xlsx file first.';
    return;
  }
  run(imp, async (client, add) => {
    const sheets = await XlsxLite.readXlsx(new Uint8Array(await f.arrayBuffer()));
    const sheet = sheets.find((x) => x.name.toLowerCase() === 'collections') || sheets[0];
    if (!sheet) throw new Error('The workbook has no sheets');
    const { collections, warnings } = Convert.rowsToCollections(sheet.rows);
    add(`Read ${collections.length} collections from sheet "${sheet.name}"`);
    warnings.forEach((w) => add(`! ${w}`));
    const s = await runImport(client, { collections }, {
      existing: document.getElementById('existing').value,
      publishAll: document.getElementById('publish').checked,
    }, add);
    add(`\nDone. Created ${s.created}, updated ${s.updated}, skipped ${s.skipped}, failed ${s.failed}, products not found ${s.missingProducts}.`);
  });
};

const shopParam = new URLSearchParams(location.search).get('shop');
if (shopParam) document.getElementById('shop').textContent = shopParam;
if (!shopParam || typeof shopify === 'undefined') document.getElementById('warn').style.display = 'block';
})();
