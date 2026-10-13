// An in-memory stand-in for the helper, used when the page is opened with `?mock=1`.
// It holds a small "Sample cube" (ride-hailing trips) and answers the same routes with
// the same shapes as the real helper, so the pane can be exercised with no helper and
// no server. The numbers are made up but consistent: totals are real sums, ratios are
// real ratios, and row/column headers follow the requested layout, including
// expand/collapse flags, totals, filters and empty-row suppression.

import type { Api } from "../api";
import { ApiFailure } from "../apiErrors";
import type {
  CellValue,
  Connection,
  Dimension,
  ErrorCategory,
  Fields,
  Grid,
  HeaderCell,
  Hierarchy,
  Layout,
  LayoutOptions,
  LineKind,
  Measure,
  Member,
  MemberRef,
  MembersRequest,
  MembersResponse,
  NewConnection,
  PlacedField,
  QueryRequest,
} from "../types";

export interface MockOptions {
  /** Make every query take four seconds, to try the spinner and Cancel. */
  slow?: boolean;
  /** Override all delays (tests use 0). */
  delayMs?: number;
}

const CATALOG = "Sample";
const CUBE = "Rides";
const CELL_LIMIT = 300000;

// ---------------------------------------------------------------------------
// Facts
// ---------------------------------------------------------------------------

interface Fact {
  month: number; // 0..23, January 2025 first
  city: number;
  service: number;
  segment: number;
  driver: number;
  rides: number;
  fare: number;
  km: number;
  cancelled: number;
}

const YEARS = [2025, 2026];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const PROVINCES: { name: string; cities: string[] }[] = [
  { name: "Tehran", cities: ["Tehran", "Rey"] },
  { name: "Fars", cities: ["Shiraz", "Marvdasht"] },
  { name: "Isfahan", cities: ["Isfahan", "Kashan"] },
  { name: "Razavi Khorasan", cities: ["Mashhad", "Neyshabur"] },
];
const CITIES: { name: string; province: number; weight: number; small: boolean }[] = [];
PROVINCES.forEach((province, p) =>
  province.cities.forEach((name, i) =>
    CITIES.push({ name, province: p, weight: i === 0 ? 9 - p * 1.5 : 1.5, small: i > 0 }),
  ),
);
const SERVICES = ["Ride", "Bike", "Box"];
const SEGMENTS = ["New", "Returning"];
const DRIVER_COUNT = 1200;

function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function buildFacts(): Fact[] {
  const random = mulberry32(47613);
  const facts: Fact[] = [];
  let index = 0;
  for (let month = 0; month < 24; month++) {
    for (let city = 0; city < CITIES.length; city++) {
      for (let service = 0; service < SERVICES.length; service++) {
        for (let segment = 0; segment < SEGMENTS.length; segment++) {
          index++;
          const noise = 0.75 + random() * 0.5;
          // Box is not offered in small cities, and Bike started in July 2025 there:
          // these gaps give "Hide empty" something to hide.
          if (CITIES[city].small && service === 2) continue;
          if (CITIES[city].small && service === 1 && month < 6) continue;
          const serviceWeight = [1, 0.35, 0.12][service];
          const segmentWeight = [0.3, 0.7][segment];
          const growth = 1 + month * 0.025;
          const rides = Math.max(
            1,
            Math.round(4200 * CITIES[city].weight * serviceWeight * segmentWeight * growth * noise),
          );
          const farePerRide = [61000, 28000, 45000][service] * (0.9 + random() * 0.2);
          facts.push({
            month,
            city,
            service,
            segment,
            driver: (index * 7919) % DRIVER_COUNT,
            rides,
            fare: Math.round(rides * farePerRide),
            km: Math.round(rides * (4 + random() * 5) * 10) / 10,
            cancelled: Math.round(rides * (0.04 + random() * 0.08)),
          });
        }
      }
    }
  }
  return facts;
}

// ---------------------------------------------------------------------------
// Dimensions
// ---------------------------------------------------------------------------

interface Node {
  uniqueName: string;
  caption: string;
  level: number;
  parent: Node | null;
  children: Node[];
  test: (fact: Fact) => boolean;
}

interface Hier {
  info: Hierarchy;
  all: Node | null;
  /** Level-0 members: the All member when there is one. */
  top: Node[];
  nodes: Map<string, Node>;
}

function makeHier(
  uniqueName: string,
  caption: string,
  kind: Hierarchy["kind"],
  hasAll: boolean,
  fill: (add: (parent: Node | null, key: string, caption: string, test: Node["test"]) => Node) => void,
  folder = "",
): Hier {
  const nodes = new Map<string, Node>();
  const top: Node[] = [];
  const all: Node | null = hasAll
    ? {
        uniqueName: `${uniqueName}.[All]`,
        caption: "All",
        level: 0,
        parent: null,
        children: [],
        test: () => true,
      }
    : null;
  if (all) {
    nodes.set(all.uniqueName, all);
    top.push(all);
  }
  fill((parent, key, text, test) => {
    const owner = parent ?? all;
    const node: Node = {
      uniqueName: `${uniqueName}.${key}`,
      caption: text,
      level: owner ? owner.level + 1 : 0,
      parent: owner,
      children: [],
      test,
    };
    if (owner) owner.children.push(node);
    else top.push(node);
    nodes.set(node.uniqueName, node);
    return node;
  });
  return {
    info: { uniqueName, caption, folder, kind, allMember: all ? all.uniqueName : null },
    all,
    top,
    nodes,
  };
}

function buildHierarchies(): { dimensions: Dimension[]; hiers: Map<string, Hier> } {
  const calendar = makeHier("[Date].[Calendar]", "Calendar", "user", true, (add) => {
    YEARS.forEach((year, y) => {
      const yearNode = add(null, `[Year].&[${year}]`, String(year), (f) => Math.floor(f.month / 12) === y);
      for (let q = 0; q < 4; q++) {
        const quarter = add(
          yearNode,
          `[Quarter].&[${year}]&[${q + 1}]`,
          `Q${q + 1} ${year}`,
          (f) => Math.floor(f.month / 3) === y * 4 + q,
        );
        for (let m = q * 3; m < q * 3 + 3; m++) {
          add(quarter, `[Month].&[${year}]&[${m + 1}]`, `${MONTH_NAMES[m]} ${year}`, (f) => f.month === y * 12 + m);
        }
      }
    });
  });
  const year = makeHier("[Date].[Year]", "Year", "attribute", true, (add) => {
    YEARS.forEach((value, y) => add(null, `&[${value}]`, String(value), (f) => Math.floor(f.month / 12) === y));
  });
  const monthName = makeHier("[Date].[Month Name]", "Month Name", "attribute", true, (add) => {
    MONTH_NAMES.forEach((name, m) => add(null, `&[${m + 1}]`, name, (f) => f.month % 12 === m));
  });
  const geography = makeHier("[Geography].[Geography]", "Geography", "user", true, (add) => {
    PROVINCES.forEach((province, p) => {
      const node = add(null, `[Province].&[${p + 1}]`, province.name, (f) => CITIES[f.city].province === p);
      CITIES.forEach((city, c) => {
        if (city.province === p) add(node, `[City].&[${c + 1}]`, city.name, (f) => f.city === c);
      });
    });
  });
  const city = makeHier("[Geography].[City]", "City", "attribute", true, (add) => {
    CITIES.forEach((entry, c) => add(null, `&[${c + 1}]`, entry.name, (f) => f.city === c));
  });
  const service = makeHier("[Service].[Service Type]", "Service Type", "attribute", true, (add) => {
    SERVICES.forEach((name, s) => add(null, `&[${s + 1}]`, name, (f) => f.service === s));
  });
  // No All member, to exercise that path: totals for it cannot be built.
  const segment = makeHier("[Customer].[Segment]", "Segment", "attribute", false, (add) => {
    SEGMENTS.forEach((name, s) => add(null, `&[${s + 1}]`, name, (f) => f.segment === s));
  });
  // Long enough to be truncated by the member picker's default limit of 500.
  const driver = makeHier(
    "[Driver].[Driver]",
    "Driver",
    "attribute",
    true,
    (add) => {
      for (let d = 0; d < DRIVER_COUNT; d++) {
        const label = "Driver " + String(d + 1).padStart(4, "0");
        add(null, `&[${d + 1}]`, label, (f) => f.driver === d);
      }
    },
    "People",
  );

  const dimension = (uniqueName: string, caption: string, list: Hier[]): Dimension => ({
    uniqueName,
    caption,
    hierarchies: list.map((h) => h.info),
  });
  const dimensions = [
    dimension("[Date]", "Date", [calendar, year, monthName]),
    dimension("[Geography]", "Geography", [geography, city]),
    dimension("[Service]", "Service", [service]),
    dimension("[Customer]", "Customer", [segment]),
    dimension("[Driver]", "Driver", [driver]),
  ];
  const hiers = new Map<string, Hier>();
  for (const h of [calendar, year, monthName, geography, city, service, segment, driver]) {
    hiers.set(h.info.uniqueName, h);
  }
  return { dimensions, hiers };
}

// ---------------------------------------------------------------------------
// Measures
// ---------------------------------------------------------------------------

interface Totals {
  count: number;
  rides: number;
  fare: number;
  km: number;
  cancelled: number;
}

interface MeasureDef {
  info: Measure;
  format: string;
  compute: (t: Totals) => number | null;
}

const MEASURES: MeasureDef[] = [
  {
    info: { uniqueName: "[Measures].[Rides]", caption: "Rides", group: "Rides", folder: "" },
    format: "#,##0",
    compute: (t) => t.rides,
  },
  {
    info: { uniqueName: "[Measures].[Cancelled Rides]", caption: "Cancelled Rides", group: "Rides", folder: "" },
    format: "#,##0",
    compute: (t) => t.cancelled,
  },
  {
    info: { uniqueName: "[Measures].[Cancel Rate]", caption: "Cancel Rate", group: "Rides", folder: "Quality" },
    format: "0.0%",
    compute: (t) => (t.rides > 0 ? t.cancelled / t.rides : null),
  },
  {
    info: { uniqueName: "[Measures].[Fare]", caption: "Fare", group: "Rides", folder: "Money" },
    format: "#,##0",
    compute: (t) => t.fare,
  },
  {
    info: { uniqueName: "[Measures].[Avg Fare]", caption: "Avg Fare", group: "Rides", folder: "Money" },
    format: "#,##0.00",
    compute: (t) => (t.rides > 0 ? t.fare / t.rides : null),
  },
  {
    info: { uniqueName: "[Measures].[Distance km]", caption: "Distance km", group: "Trips", folder: "" },
    format: "#,##0.0",
    compute: (t) => t.km,
  },
  {
    info: { uniqueName: "[Measures].[Avg Distance km]", caption: "Avg Distance km", group: "Trips", folder: "" },
    format: "0.00",
    compute: (t) => (t.rides > 0 ? t.km / t.rides : null),
  },
];

function totalsOf(facts: Fact[]): Totals {
  const t: Totals = { count: facts.length, rides: 0, fare: 0, km: 0, cancelled: 0 };
  for (const f of facts) {
    t.rides += f.rides;
    t.fare += f.fare;
    t.km += f.km;
    t.cancelled += f.cancelled;
  }
  return t;
}

// ---------------------------------------------------------------------------
// Query
// ---------------------------------------------------------------------------

interface Tuple {
  nodes: Node[];
  /** Index where the All members start; equal to nodes.length for a detail tuple. */
  totalFrom: number;
  kind: LineKind;
}

function fail(category: ErrorCategory, status: number, message: string, detail?: string): ApiFailure {
  return new ApiFailure({ category, message, detail }, status);
}

function isAncestor(ancestor: Node, node: Node): boolean {
  for (let p = node.parent; p; p = p.parent) if (p === ancestor) return true;
  return false;
}

class SampleCube {
  readonly facts = buildFacts();
  readonly dimensions: Dimension[];
  readonly hiers: Map<string, Hier>;

  constructor() {
    const built = buildHierarchies();
    this.dimensions = built.dimensions;
    this.hiers = built.hiers;
  }

  fields(): Fields {
    return { measures: MEASURES.map((m) => m.info), dimensions: this.dimensions };
  }

  hier(uniqueName: string): Hier {
    const found = this.hiers.get(uniqueName);
    if (!found) {
      throw fail("not-found", 404, "This cube has no such field. It may have been renamed.", uniqueName);
    }
    return found;
  }

  members(request: MembersRequest): MembersResponse {
    const h = this.hier(request.hierarchy);
    const limit = Math.min(5000, Math.max(1, Math.floor(request.limit ?? 500)));
    let list: Node[];
    if (request.search !== undefined && request.search.trim() !== "") {
      const needle = request.search.trim().toLowerCase();
      list = Array.from(h.nodes.values()).filter(
        (n) => n !== h.all && n.caption.toLowerCase().includes(needle),
      );
    } else if (request.parent !== undefined) {
      const parent = h.nodes.get(request.parent);
      if (!parent) throw fail("not-found", 404, "That item is no longer in the cube.", request.parent);
      list = parent.children;
    } else {
      list = h.top;
    }
    return {
      members: list.slice(0, limit).map(
        (n): Member => ({
          uniqueName: n.uniqueName,
          caption: n.caption,
          level: n.level,
          hasChildren: n.children.length > 0,
        }),
      ),
      truncated: list.length > limit,
    };
  }

  private selected(h: Hier, field: PlacedField): Node[] {
    const out: Node[] = [];
    for (const name of field.members) {
      const node = h.nodes.get(name);
      if (node) out.push(node);
    }
    return out;
  }

  /** The detail set of one axis field: top level, drilled down where expanded. */
  private detail(h: Hier, field: PlacedField): Node[] {
    const selected = this.selected(h, field);
    const visible = (n: Node) =>
      selected.length === 0 || selected.some((s) => s === n || isAncestor(n, s) || isAncestor(s, n));
    const expanded = new Set(field.expanded);
    const out: Node[] = [];
    const walk = (nodes: Node[]) => {
      for (const n of nodes) {
        if (!visible(n)) continue;
        out.push(n);
        if (expanded.has(n.uniqueName)) walk(n.children);
      }
    };
    walk(h.all ? h.all.children : h.top);
    return out;
  }

  private axis(fields: PlacedField[], options: LayoutOptions): Tuple[] {
    const n = fields.length;
    if (n === 0) return [{ nodes: [], totalFrom: 0, kind: "data" }];
    const hiers = fields.map((f) => this.hier(f.hierarchy));
    const details = fields.map((f, i) => this.detail(hiers[i], f));

    const build = (k: number): { nodes: Node[]; totalFrom: number }[] => {
      const out: { nodes: Node[]; totalFrom: number }[] = [];
      const tail = k + 1 < n ? build(k + 1) : null;
      for (const d of details[k]) {
        if (tail) for (const t of tail) out.push({ nodes: [d, ...t.nodes], totalFrom: t.totalFrom });
        else out.push({ nodes: [d], totalFrom: n });
      }
      const wanted = k === 0 ? options.grandTotals : options.subtotals;
      const rest = hiers.slice(k);
      if (wanted && rest.every((h) => h.all !== null)) {
        out.push({ nodes: rest.map((h) => h.all as Node), totalFrom: k });
      }
      return out;
    };

    return build(0).map((t) => ({
      nodes: t.nodes,
      totalFrom: t.totalFrom,
      kind: t.totalFrom === n ? "data" : t.totalFrom === 0 ? "grandTotal" : "subtotal",
    }));
  }

  private memberRef(node: Node, field: PlacedField): MemberRef {
    return {
      uniqueName: node.uniqueName,
      hierarchy: field.hierarchy,
      caption: node.caption,
      level: node.level,
      hasChildren: node.children.length > 0,
      expanded: field.expanded.includes(node.uniqueName),
    };
  }

  /**
   * The header cell for position `r` of `tuple`. `previous` is the tuple before it on
   * the axis (or the same tuple, for a repeated measure column), used to blank repeats.
   */
  private headerCell(
    tuple: Tuple,
    r: number,
    previous: Tuple | null,
    fields: PlacedField[],
    hiers: Hier[],
  ): HeaderCell {
    const isMember = (t: Tuple, at: number) => (t.kind === "data" ? true : at < t.totalFrom - 1);
    if (!isMember(tuple, r)) {
      const kind = tuple.kind === "grandTotal" ? "grandTotal" : "subtotal";
      const labelAt = Math.max(0, tuple.totalFrom - 1);
      let text = "";
      if (r === labelAt && previous !== tuple) {
        text = tuple.kind === "grandTotal" ? "Grand Total" : `${tuple.nodes[r].caption} Total`;
      }
      return { text, indent: 0, kind, member: null };
    }
    const node = tuple.nodes[r];
    const member = this.memberRef(node, fields[r]);
    const repeated =
      previous !== null &&
      isMember(previous, r) &&
      previous.nodes.slice(0, r + 1).every((n, i) => n === tuple.nodes[i]);
    if (repeated) return { text: "", indent: 0, kind: "blank", member };
    const base = hiers[r].all ? 1 : 0;
    return { text: node.caption, indent: Math.max(0, node.level - base), kind: "member", member };
  }

  query(layout: Layout, elapsedMs: number): Grid {
    const rowHiers = layout.rows.map((f) => this.hier(f.hierarchy));
    const colHiers = layout.columns.map((f) => this.hier(f.hierarchy));

    const measures =
      layout.values.length > 0
        ? layout.values.map((v) => {
            const def = MEASURES.find((m) => m.info.uniqueName === v.measure);
            if (!def) {
              throw fail("not-found", 404, "This cube has no such measure. It may have been renamed.", v.measure);
            }
            return { def, caption: v.caption };
          })
        : [{ def: MEASURES[0], caption: MEASURES[0].info.caption }];

    // Every field with selected members restricts the facts, whatever area it is in.
    const restrictions: Node[][] = [];
    for (const field of [...layout.rows, ...layout.columns, ...layout.filters]) {
      const selected = this.selected(this.hier(field.hierarchy), field);
      if (selected.length > 0) restrictions.push(selected);
    }
    const base = this.facts.filter((f) => restrictions.every((set) => set.some((n) => n.test(f))));

    let rowTuples = this.axis(layout.rows, layout.options);
    let colTuples = this.axis(layout.columns, layout.options);
    const cellCount = rowTuples.length * colTuples.length * measures.length;
    if (cellCount > CELL_LIMIT) {
      throw fail(
        "too-large",
        413,
        `This view would return ${cellCount.toLocaleString("en-US")} cells. Add a filter or remove a field.`,
      );
    }

    const matches = (tuple: Tuple, facts: Fact[]) =>
      tuple.nodes.length === 0 ? facts : facts.filter((f) => tuple.nodes.every((n) => n.test(f)));

    // totals[row][col]
    let totals: Totals[][] = rowTuples.map((rowTuple) => {
      const rowFacts = matches(rowTuple, base);
      return colTuples.map((colTuple) => totalsOf(matches(colTuple, rowFacts)));
    });

    if (layout.options.nonEmpty) {
      const keepRow = rowTuples.map((_, r) => totals[r].some((t) => t.count > 0));
      const keepCol = colTuples.map((_, c) => totals.some((line) => line[c].count > 0));
      rowTuples = rowTuples.filter((_, r) => keepRow[r]);
      colTuples = colTuples.filter((_, c) => keepCol[c]);
      totals = totals
        .filter((_, r) => keepRow[r])
        .map((line) => line.filter((_, c) => keepCol[c]));
    }

    const formatPalette = ["General"];
    const formatIndex = measures.map((m) => {
      let index = formatPalette.indexOf(m.def.format);
      if (index < 0) index = formatPalette.push(m.def.format) - 1;
      return index;
    });

    // Columns: every column tuple once per measure, measures innermost.
    const colHeaderRows: HeaderCell[][] = layout.columns.map(() => []);
    const measureRow: HeaderCell[] = [];
    const colKinds: LineKind[] = [];
    colTuples.forEach((tuple, c) => {
      measures.forEach((measure, m) => {
        const previous = m > 0 ? tuple : c > 0 ? colTuples[c - 1] : null;
        layout.columns.forEach((_, r) => {
          colHeaderRows[r].push(this.headerCell(tuple, r, previous, layout.columns, colHiers));
        });
        measureRow.push({ text: measure.caption, indent: 0, kind: "measure", member: null });
        colKinds.push(tuple.kind);
      });
    });
    colHeaderRows.push(measureRow);

    // Rows.
    const hasRowFields = layout.rows.length > 0;
    const rowHeaders: HeaderCell[][] = rowTuples.map((tuple, r) =>
      hasRowFields
        ? layout.rows.map((_, i) =>
            this.headerCell(tuple, i, r > 0 ? rowTuples[r - 1] : null, layout.rows, rowHiers),
          )
        : [{ text: "Total", indent: 0, kind: "grandTotal", member: null }],
    );
    const rowKinds: LineKind[] = rowTuples.map((t) => t.kind);

    const values: CellValue[][] = [];
    const cellFormats: number[][] = [];
    totals.forEach((line) => {
      const out: CellValue[] = [];
      const formats: number[] = [];
      line.forEach((t) => {
        measures.forEach((measure, m) => {
          out.push(t.count > 0 ? measure.def.compute(t) : null);
          formats.push(formatIndex[m]);
        });
      });
      values.push(out);
      cellFormats.push(formats);
    });

    const filterRows = layout.filters.map((field) => ({
      caption: field.caption,
      value:
        field.members.length === 0
          ? "All"
          : field.members.length === 1
            ? field.memberCaptions[0] ?? field.members[0]
            : "(Multiple Items)",
    }));

    return {
      filterRows,
      rowFieldCaptions: hasRowFields ? layout.rows.map((f) => f.caption) : [""],
      colHeaderRows,
      rowHeaders,
      values,
      formatPalette,
      cellFormats,
      rowKinds,
      colKinds,
      stats: {
        rows: rowHeaders.length,
        columns: colKinds.length,
        elapsedMs,
        mdx: sampleMdx(layout),
      },
    };
  }
}

// ---------------------------------------------------------------------------
// A plausible MDX string, for "Show query". Follows the helper's rules loosely.
// ---------------------------------------------------------------------------

function sampleMdx(layout: Layout): string {
  const detail = (f: PlacedField) => {
    const set = f.allMember ? `{${f.allMember}}.Children` : `${f.hierarchy}.Levels(0).Members`;
    return f.expanded.length > 0
      ? `DrilldownMember(${set}, {${f.expanded.join(", ")}}, RECURSIVE)`
      : set;
  };
  const axisSet = (fields: PlacedField[], options: LayoutOptions): string => {
    const build = (k: number): string => {
      const rest = fields.slice(k);
      const inner = k + 1 < fields.length ? `CrossJoin(${detail(fields[k])}, ${build(k + 1)})` : detail(fields[k]);
      const wanted = k === 0 ? options.grandTotals : options.subtotals;
      if (!wanted || rest.some((f) => f.allMember === null)) return inner;
      const all = rest.map((f) => f.allMember).join(", ");
      return `{${inner}, {${rest.length > 1 ? `(${all})` : all}}}`;
    };
    return build(0);
  };
  const measures = `{${
    layout.values.length > 0 ? layout.values.map((v) => v.measure).join(", ") : "[Measures].DefaultMember"
  }}`;
  const nonEmpty = layout.options.nonEmpty ? "NON EMPTY " : "";
  const columns =
    layout.columns.length > 0
      ? `CrossJoin(${axisSet(layout.columns, layout.options)}, ${measures})`
      : measures;

  let from = `[${CUBE}]`;
  const where: string[] = [];
  for (const field of [...layout.rows, ...layout.columns]) {
    if (field.members.length > 0) from = `(SELECT ({${field.members.join(", ")}}) ON COLUMNS FROM ${from})`;
  }
  for (const field of layout.filters) {
    if (field.members.length === 1) where.push(field.members[0]);
    else if (field.members.length > 1) {
      from = `(SELECT ({${field.members.join(", ")}}) ON COLUMNS FROM ${from})`;
    }
  }

  const lines = [`SELECT`, `  ${nonEmpty}${columns} ON COLUMNS`];
  if (layout.rows.length > 0) {
    lines[1] += ",";
    lines.push(`  ${nonEmpty}${axisSet(layout.rows, layout.options)} ON ROWS`);
  }
  lines.push(`FROM ${from}`);
  if (where.length > 0) lines.push(`WHERE (${where.join(", ")})`);
  lines.push(`CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE`);
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// The client
// ---------------------------------------------------------------------------

export function createMockApi(options: MockOptions = {}): Api {
  const cube = new SampleCube();
  const connections: Connection[] = [
    { id: "demo", name: "Sample cube", kind: "demo", server: "", username: "" },
  ];
  let nextId = 1;
  /** Queries in flight, per connection, so that `cancel` can stop them. */
  const running = new Map<string, Set<(error: ApiFailure) => void>>();

  const delay = (ms: number) => (options.delayMs !== undefined ? options.delayMs : ms);
  const wait = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, delay(ms)));

  const need = (connectionId: string) => {
    const found = connections.find((c) => c.id === connectionId);
    if (!found) {
      throw fail("not-found", 404, "This connection no longer exists.", connectionId);
    }
    return found;
  };
  const needCube = (catalog: string, name: string) => {
    if (catalog !== CATALOG) throw fail("not-found", 404, `There is no database named "${catalog}".`);
    if (name !== CUBE) throw fail("not-found", 404, `There is no cube named "${name}".`);
  };

  return {
    async status() {
      await wait(20);
      return { ok: true, version: "0.1.0-mock", certDaysLeft: 800 };
    },

    async connections() {
      await wait(40);
      return connections.map((c) => ({ ...c }));
    },

    async addConnection(input: NewConnection) {
      await wait(600);
      const server = input.server.trim();
      if (server === "" || /[\s,]/.test(server)) {
        throw fail("bad-request", 400, "Enter the server as host, host:port or host\\instance.");
      }
      if (input.username.trim() === "") {
        throw fail("bad-request", 400, "Enter the username as DOMAIN\\user.");
      }
      if (/down|offline/i.test(server)) {
        throw fail(
          "unreachable",
          502,
          "Can't reach the server. Are you on the VPN?",
          `connect ${server}:2383: timed out`,
        );
      }
      if (input.password === "" || input.password === "wrong") {
        throw fail(
          "auth",
          401,
          "The server refused the sign-in. Check username (DOMAIN\\user) and password.",
          "NTLM: STATUS_LOGON_FAILURE",
        );
      }
      const connection: Connection = {
        id: `mock-${nextId++}`,
        name: (input.name || "").trim() || server,
        kind: "ssas",
        server,
        username: input.username.trim(),
      };
      connections.push(connection);
      return { ...connection };
    },

    async deleteConnection(connectionId) {
      await wait(60);
      const found = need(connectionId);
      if (found.kind === "demo") {
        throw fail("bad-request", 400, "The Sample cube cannot be deleted.");
      }
      connections.splice(connections.indexOf(found), 1);
    },

    async catalogs(connectionId) {
      await wait(80);
      need(connectionId);
      return [{ name: CATALOG, description: "Made-up ride data for trying the add-in" }];
    },

    async cubes(connectionId, catalog) {
      await wait(80);
      need(connectionId);
      if (catalog !== CATALOG) throw fail("not-found", 404, `There is no database named "${catalog}".`);
      return [{ name: CUBE, caption: "Rides" }];
    },

    async fields(connectionId, catalog, name) {
      await wait(120);
      need(connectionId);
      needCube(catalog, name);
      return cube.fields();
    },

    async members(connectionId, request) {
      await wait(120);
      need(connectionId);
      needCube(request.catalog, request.cube);
      return cube.members(request);
    },

    query(connectionId, request: QueryRequest, signal) {
      const started = Date.now();
      return new Promise<Grid>((resolve, reject) => {
        const set = running.get(connectionId) ?? new Set<(error: ApiFailure) => void>();
        running.set(connectionId, set);
        let settled = false;
        let timer: ReturnType<typeof setTimeout> | undefined;
        const stop = (error: ApiFailure) => {
          if (settled) return;
          settled = true;
          if (timer !== undefined) clearTimeout(timer);
          set.delete(stop);
          reject(error);
        };
        set.add(stop);
        if (signal) {
          const onAbort = () => stop(fail("cancelled", 499, "Cancelled."));
          if (signal.aborted) {
            onAbort();
            return;
          }
          signal.addEventListener("abort", onAbort);
        }
        timer = setTimeout(
          () => {
            if (settled) return;
            settled = true;
            set.delete(stop);
            try {
              need(connectionId);
              needCube(request.catalog, request.cube);
              resolve(cube.query(request.layout, Date.now() - started));
            } catch (error) {
              reject(error);
            }
          },
          delay(options.slow ? 4000 : 220),
        );
      });
    },

    async cancel(connectionId) {
      const set = running.get(connectionId);
      if (set) {
        for (const stop of Array.from(set)) stop(fail("cancelled", 499, "The query was cancelled."));
      }
    },
  };
}
