// An in-memory stand-in for the helper, used when the page is opened with `?mock=1`.
// It holds a small "Sample cube" (ride-hailing trips) and answers the same routes with
// the same shapes as the real helper, so the pane can be exercised with no helper and
// no server. The numbers are made up but consistent: totals are real sums, ratios are
// real ratios, and row/column headers follow the requested layout, including
// expand/collapse flags, totals, filters and empty-row suppression. Sorting, every
// label and value filter, Show Values As and Show Details are computed for real on the
// sample facts: a filter the mock cannot apply would be an error, never a no-op.

import type { Api } from "../api";
import { ApiFailure } from "../apiErrors";
import type {
  CellValue,
  Connection,
  DetailsRequest,
  DetailsResponse,
  Dimension,
  ErrorCategory,
  Fields,
  Grid,
  HeaderCell,
  Hierarchy,
  LabelFilter,
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
  ShowValuesAs,
  ValueFilter,
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
      // As the real helper does: the top level is what sits under the All member.
      // The All member itself is never listed.
      list = h.all ? h.all.children : h.top;
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

  private measureDef(name: string): MeasureDef {
    const def = MEASURES.find((m) => m.info.uniqueName === name);
    if (!def) throw fail("not-found", 404, "This cube has no such measure. It may have been renamed.", name);
    return def;
  }

  private nodeCaption(name: string): string {
    for (const h of this.hiers.values()) {
      const node = h.nodes.get(name);
      if (node) return node.caption;
    }
    return name;
  }

  private findNode(name: string): Node {
    for (const h of this.hiers.values()) {
      const node = h.nodes.get(name);
      if (node) return node;
    }
    throw fail("not-found", 404, "That item is no longer in the cube.", name);
  }

  /** A field's top-level items that its item selection leaves visible. */
  private topItems(h: Hier, field: PlacedField): Node[] {
    const selected = this.selected(h, field);
    const top = h.all ? h.all.children : h.top;
    return top.filter(
      (n) => selected.length === 0 || selected.some((s) => s === n || isAncestor(n, s) || isAncestor(s, n)),
    );
  }

  /**
   * The top-level items a row or column field keeps after its label and value filters,
   * or null when it has neither. Values are measured over `facts`.
   */
  private kept(h: Hier, field: PlacedField, facts: Fact[]): Set<Node> | null {
    const label = field.labelFilter;
    const value = field.valueFilter;
    if (!label && !value) return null;
    let items = this.topItems(h, field);
    if (label) items = items.filter((n) => labelMatches(label, n.caption));
    if (value) {
      const def = this.measureDef(value.measure);
      items = valueFiltered(value, items, (n) => valueOf(def, facts.filter(n.test)));
    }
    return new Set(items);
  }

  /**
   * The facts a layout looks at: every field's item selection, then the label and value
   * filters of the row and column fields (so totals count only the items left).
   */
  private restricted(layout: Layout): { base: Fact[]; keptRows: (Set<Node> | null)[]; keptCols: (Set<Node> | null)[] } {
    const restrictions: Node[][] = [];
    for (const field of [...layout.rows, ...layout.columns, ...layout.filters]) {
      const selected = this.selected(this.hier(field.hierarchy), field);
      if (selected.length > 0) restrictions.push(selected);
    }
    const selectedFacts = this.facts.filter((f) => restrictions.every((set) => set.some((n) => n.test(f))));
    const keptRows = layout.rows.map((f) => this.kept(this.hier(f.hierarchy), f, selectedFacts));
    const keptCols = layout.columns.map((f) => this.kept(this.hier(f.hierarchy), f, selectedFacts));
    const narrowing = [...keptRows, ...keptCols]
      .filter((set): set is Set<Node> => set !== null)
      .map((set) => Array.from(set));
    const base = selectedFacts.filter((f) => narrowing.every((list) => list.some((n) => n.test(f))));
    return { base, keptRows, keptCols };
  }

  /** The detail set of one axis field: top level, drilled down where expanded, in its sort order. */
  private detail(h: Hier, field: PlacedField, kept: Set<Node> | null, facts: Fact[]): Node[] {
    const selected = this.selected(h, field);
    const top = new Set(h.all ? h.all.children : h.top);
    const visible = (n: Node) =>
      (selected.length === 0 || selected.some((s) => s === n || isAncestor(n, s) || isAncestor(s, n))) &&
      (kept === null || !top.has(n) || kept.has(n));
    const sort = field.sort;
    const order = (nodes: Node[]): Node[] => {
      if (!sort) return nodes;
      const sign = sort.order === "asc" ? 1 : -1;
      if (sort.by === "caption") {
        return nodes
          .slice()
          .sort((a, b) => sign * a.caption.localeCompare(b.caption, "en", { numeric: true, sensitivity: "base" }));
      }
      const def = this.measureDef(sort.measure ?? "");
      const scored = nodes.map((n) => ({ n, v: valueOf(def, facts.filter(n.test)) }));
      scored.sort((a, b) => {
        if (a.v === null || b.v === null) return a.v === null ? (b.v === null ? 0 : 1) : -1;
        return sign * (a.v - b.v);
      });
      return scored.map((s) => s.n);
    };
    const expanded = new Set(field.expanded);
    const out: Node[] = [];
    const walk = (nodes: Node[]) => {
      for (const n of order(nodes)) {
        if (!visible(n)) continue;
        out.push(n);
        if (expanded.has(n.uniqueName)) walk(n.children);
      }
    };
    walk(Array.from(top));
    return out;
  }

  private axis(
    fields: PlacedField[],
    options: LayoutOptions,
    grandTotal: boolean,
    kept: (Set<Node> | null)[],
    facts: Fact[],
  ): Tuple[] {
    const n = fields.length;
    if (n === 0) return [{ nodes: [], totalFrom: 0, kind: "data" }];
    const hiers = fields.map((f) => this.hier(f.hierarchy));
    const details = fields.map((f, i) => this.detail(hiers[i], f, kept[i] ?? null, facts));

    const build = (k: number): { nodes: Node[]; totalFrom: number }[] => {
      const out: { nodes: Node[]; totalFrom: number }[] = [];
      const tail = k + 1 < n ? build(k + 1) : null;
      for (const d of details[k]) {
        if (tail) for (const t of tail) out.push({ nodes: [d, ...t.nodes], totalFrom: t.totalFrom });
        else out.push({ nodes: [d], totalFrom: n });
      }
      const wanted = k === 0 ? grandTotal : options.subtotals;
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
            const showAs = v.showAs ?? "normal";
            const label = SHOW_AS_CAPTIONS[showAs];
            // As the helper words it: "Rides (Difference From City, previous)".
            const field = v.baseField ? lastBracketed(v.baseField) : "";
            const item =
              v.baseItem === "previous" || v.baseItem === "next"
                ? v.baseItem
                : v.baseItem
                  ? this.nodeCaption(v.baseItem)
                  : "";
            const detail = (field ? " " + field : "") + (item ? ", " + item : "");
            return {
              def: this.measureDef(v.measure),
              caption: label ? `${v.caption} (${label}${detail})` : v.caption,
              showAs,
              baseField: v.baseField ?? null,
              baseItem: v.baseItem ?? null,
            };
          })
        : [
            {
              def: MEASURES[0],
              caption: MEASURES[0].info.caption,
              showAs: "normal" as ShowValuesAs,
              baseField: null as string | null,
              baseItem: null as string | null,
            },
          ];

    const { base, keptRows, keptCols } = this.restricted(layout);

    // The Grand Total row closes the rows axis; the Grand Total column, the columns axis.
    const options = layout.options;
    let rowTuples = this.axis(layout.rows, options, options.grandTotalRow ?? options.grandTotals, keptRows, base);
    let colTuples = this.axis(layout.columns, options, options.grandTotalColumn ?? options.grandTotals, keptCols, base);
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

    // "Show Values As": every kind is worked out from the facts, cell by cell.
    const shown = measures.map((measure) =>
      measure.showAs === "normal" ? null : this.showAs(measure, layout, base),
    );

    const formatPalette = ["General"];
    const formatIndex = measures.map((m) => {
      const format = SHOW_AS_FORMATS[m.showAs] ?? m.def.format;
      let index = formatPalette.indexOf(format);
      if (index < 0) index = formatPalette.push(format) - 1;
      return index;
    });

    // Columns: every column tuple once per measure, measures innermost.
    const colHeaderRows: HeaderCell[][] = layout.columns.map(() => []);
    const measureRow: HeaderCell[] = [];
    const colKinds: LineKind[] = [];
    const colValueIndex: number[] = [];
    colTuples.forEach((tuple, c) => {
      measures.forEach((measure, m) => {
        colValueIndex.push(layout.values.length > 0 ? m : -1);
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
    totals.forEach((line, r) => {
      const out: CellValue[] = [];
      const formats: number[] = [];
      line.forEach((t, c) => {
        measures.forEach((measure, m) => {
          const raw = t.count > 0 ? measure.def.compute(t) : null;
          const calc = shown[m];
          out.push(calc && raw !== null ? calc(rowTuples[r].nodes, colTuples[c].nodes, raw) : raw);
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
          ? field.allMember !== null
            ? "All"
            : "(Default)"
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
      colValueIndex,
      stats: {
        rows: rowHeaders.length,
        columns: colKinds.length,
        elapsedMs,
        mdx: sampleMdx(layout),
      },
    };
  }

  /**
   * One Show Values As kind as a function of a cell's row nodes, column nodes and own
   * value. Every other value it needs is computed from the same facts, so the sample
   * numbers obey the same arithmetic the cube's would.
   */
  private showAs(
    measure: { def: MeasureDef; showAs: ShowValuesAs; baseField: string | null; baseItem: string | null },
    layout: Layout,
    facts: Fact[],
  ): (rowNodes: Node[], colNodes: Node[], value: number) => number | null {
    const memo = new Map<string, number | null>();
    const at = (rowNodes: Node[], colNodes: Node[]): number | null => {
      const nodes = [...rowNodes, ...colNodes];
      const key = nodes.map((n) => n.uniqueName).join("|");
      if (memo.has(key)) return memo.get(key) as number | null;
      const result = valueOf(measure.def, facts.filter((f) => nodes.every((n) => n.test(f))));
      memo.set(key, result);
      return result;
    };
    const ratio = (a: number | null, b: number | null) => (a === null || b === null || b === 0 ? null : a / b);
    const allNodes = new Set<Node>();
    for (const h of this.hiers.values()) if (h.all) allNodes.add(h.all);
    const isAll = (n: Node) => allNodes.has(n);
    /** Its parent on its axis: the parent member, or (at the top) the outer item. */
    const parentOn = (nodes: Node[]): Node[] | null => {
      for (let i = nodes.length - 1; i >= 0; i--) {
        if (isAll(nodes[i])) continue;
        const copy = nodes.slice();
        const parent = nodes[i].parent;
        if (parent) copy[i] = parent;
        else return null;
        return copy;
      }
      return null;
    };

    switch (measure.showAs) {
      case "percentOfGrandTotal":
        return (_r, _c, v) => ratio(v, at([], []));
      case "percentOfColumnTotal":
        return (_r, c, v) => ratio(v, at([], c));
      case "percentOfRowTotal":
        return (r, _c, v) => ratio(v, at(r, []));
      case "index":
        return (r, c, v) => ratio(v * (at([], []) ?? 0), (at(r, []) ?? 0) * (at([], c) ?? 0));
      case "percentOfParentRowTotal":
        return (r, c, v) => {
          const up = parentOn(r);
          return up ? ratio(v, at(up, c)) : 1;
        };
      case "percentOfParentColumnTotal":
        return (r, c, v) => {
          const up = parentOn(c);
          return up ? ratio(v, at(r, up)) : 1;
        };
      default:
        break;
    }

    // The kinds that work along a base field.
    const onRows = layout.rows.findIndex((f) => f.hierarchy === measure.baseField);
    const onCols = layout.columns.findIndex((f) => f.hierarchy === measure.baseField);
    if (onRows < 0 && onCols < 0) return () => null;
    const h = this.hier(measure.baseField as string);
    const position = onRows >= 0 ? onRows : onCols;
    const pick = (r: Node[], c: Node[]) => (onRows >= 0 ? r : c)[position];
    const put = (r: Node[], c: Node[], node: Node): [Node[], Node[]] => {
      const copy = (onRows >= 0 ? r : c).slice();
      copy[position] = node;
      return onRows >= 0 ? [copy, c] : [r, copy];
    };
    const siblings = (n: Node) => (n.parent ? n.parent.children : h.top);
    const baseOf = (n: Node): Node | null => {
      const list = siblings(n);
      const i = list.indexOf(n);
      if (measure.baseItem === "previous") return i > 0 ? list[i - 1] : null;
      if (measure.baseItem === "next") return i >= 0 && i < list.length - 1 ? list[i + 1] : null;
      return h.nodes.get(measure.baseItem ?? "") ?? null;
    };
    const valueWith = (r: Node[], c: Node[], n: Node) => {
      const [rr, cc] = put(r, c, n);
      return at(rr, cc);
    };

    switch (measure.showAs) {
      case "percentOf":
        return (r, c, v) => {
          const n = pick(r, c);
          const b = n && !isAll(n) ? baseOf(n) : null;
          return b ? ratio(v, valueWith(r, c, b)) : null;
        };
      case "differenceFrom":
      case "percentDifferenceFrom":
        return (r, c, v) => {
          const n = pick(r, c);
          const b = n && !isAll(n) ? baseOf(n) : null;
          if (!b || b === n) return null;
          const other = valueWith(r, c, b);
          if (other === null) return null;
          return measure.showAs === "differenceFrom" ? v - other : ratio(v - other, other);
        };
      case "percentOfParentTotal":
        return (r, c, v) => {
          // Everything nested inside the base field, on its axis, at All.
          const own = (onRows >= 0 ? r : c).map((node, i) => {
            if (i <= position) return node;
            const field = (onRows >= 0 ? layout.rows : layout.columns)[i];
            return this.hier(field.hierarchy).all ?? node;
          });
          return ratio(v, onRows >= 0 ? at(own, c) : at(r, own));
        };
      case "runningTotalIn":
      case "percentRunningTotalIn":
        return (r, c) => {
          const n = pick(r, c);
          if (!n || isAll(n)) return null;
          const list = siblings(n);
          let running = 0;
          for (const s of list.slice(0, list.indexOf(n) + 1)) running += valueWith(r, c, s) ?? 0;
          if (measure.showAs === "runningTotalIn") return running;
          const whole = list.reduce((sum, s) => sum + (valueWith(r, c, s) ?? 0), 0);
          return ratio(running, whole);
        };
      case "rankAscending":
      case "rankDescending":
        return (r, c, v) => {
          const n = pick(r, c);
          if (!n || isAll(n)) return null;
          const others = siblings(n)
            .map((s) => valueWith(r, c, s))
            .filter((x): x is number => x !== null);
          const better = others.filter((x) => (measure.showAs === "rankAscending" ? x < v : x > v)).length;
          return better + 1;
        };
      default:
        return () => null;
    }
  }

  /** The fact rows behind one value cell, as Show Details on a PivotTable returns them. */
  details(request: DetailsRequest): DetailsResponse {
    const layout = request.layout;
    const def = this.measureDef(request.measure);
    const nodes = request.members.map((name) => this.findNode(name));
    // As the helper does: a field filtered to several items must be pinned to one of
    // them by the cell itself, or there is no single slice of facts to show.
    for (const field of [...layout.rows, ...layout.columns, ...layout.filters]) {
      if (field.members.length < 2) continue;
      const h = this.hier(field.hierarchy);
      if (nodes.some((n) => h.nodes.get(n.uniqueName) === n)) continue;
      throw fail(
        "bad-request",
        400,
        `Show Details is not available here because '${field.caption}' is filtered to several items. ` +
          "Filter it to one item, or select a cell that belongs to one of them.",
      );
    }
    const { base } = this.restricted(layout);
    const rows = base.filter((f) => nodes.every((n) => n.test(f)));
    const max = Math.min(100000, Math.max(1, Math.floor(request.maxRows ?? 1000)));
    return {
      columns: ["Month", "Province", "City", "Service Type", "Segment", "Driver", def.info.caption],
      rows: rows.slice(0, max).map((f) => [
        `${MONTH_NAMES[f.month % 12]} ${YEARS[Math.floor(f.month / 12)]}`,
        PROVINCES[CITIES[f.city].province].name,
        CITIES[f.city].name,
        SERVICES[f.service],
        SEGMENTS[f.segment],
        "Driver " + String(f.driver + 1).padStart(4, "0"),
        def.compute(totalsOf([f])),
      ]),
      truncated: rows.length > max,
    };
  }
}

const SHOW_AS_CAPTIONS: Record<ShowValuesAs, string | null> = {
  normal: null,
  percentOfGrandTotal: "% of Grand Total",
  percentOfColumnTotal: "% of Column Total",
  percentOfRowTotal: "% of Row Total",
  percentOf: "% Of",
  percentOfParentRowTotal: "% of Parent Row Total",
  percentOfParentColumnTotal: "% of Parent Column Total",
  percentOfParentTotal: "% of Parent Total",
  differenceFrom: "Difference From",
  percentDifferenceFrom: "% Difference From",
  runningTotalIn: "Running Total In",
  percentRunningTotalIn: "% Running Total In",
  rankAscending: "Rank Smallest to Largest",
  rankDescending: "Rank Largest to Smallest",
  index: "Index",
};

/** Number format of each kind; absent keeps the measure's own. */
const SHOW_AS_FORMATS: Partial<Record<ShowValuesAs, string>> = {
  percentOfGrandTotal: "0.00%",
  percentOfColumnTotal: "0.00%",
  percentOfRowTotal: "0.00%",
  percentOf: "0.00%",
  percentOfParentRowTotal: "0.00%",
  percentOfParentColumnTotal: "0.00%",
  percentOfParentTotal: "0.00%",
  percentDifferenceFrom: "0.00%",
  percentRunningTotalIn: "0.00%",
  rankAscending: "#,##0",
  rankDescending: "#,##0",
  index: "0.00",
  // As the helper: no format of their own, so a user number format matters most here.
  differenceFrom: "General",
  runningTotalIn: "General",
};

/** "[Geography].[City]" -> "City". */
function lastBracketed(uniqueName: string): string {
  const match = /\[([^\]]*)\]\s*$/.exec(uniqueName);
  return match ? match[1] : uniqueName;
}

function valueOf(def: MeasureDef, facts: Fact[]): number | null {
  const t = totalsOf(facts);
  return t.count > 0 ? def.compute(t) : null;
}

/** Labels compare without regard to case, as the contract says. */
export function labelMatches(filter: LabelFilter, caption: string): boolean {
  const c = caption.toLowerCase();
  const v = filter.value.toLowerCase();
  const v2 = (filter.value2 ?? "").toLowerCase();
  switch (filter.operator) {
    case "equals":
      return c === v;
    case "notEquals":
      return c !== v;
    case "beginsWith":
      return c.indexOf(v) === 0;
    case "notBeginsWith":
      return c.indexOf(v) !== 0;
    case "endsWith":
      return c.length >= v.length && c.slice(c.length - v.length) === v;
    case "notEndsWith":
      return !(c.length >= v.length && c.slice(c.length - v.length) === v);
    case "contains":
      return c.indexOf(v) >= 0;
    case "notContains":
      return c.indexOf(v) < 0;
    case "greaterThan":
      return c > v;
    case "greaterOrEqual":
      return c >= v;
    case "lessThan":
      return c < v;
    case "lessOrEqual":
      return c <= v;
    case "between":
      return c >= v && c <= v2;
    case "notBetween":
      return !(c >= v && c <= v2);
  }
}

/** The items a value filter keeps. Items with no value are never kept. */
export function valueFiltered<T>(filter: ValueFilter, items: T[], value: (item: T) => number | null): T[] {
  const scored = items
    .map((item) => ({ item, v: value(item) }))
    .filter((s): s is { item: T; v: number } => s.v !== null);
  const same = (a: number, b: number) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));
  const x = filter.value;
  const y = filter.value2 ?? filter.value;
  switch (filter.operator) {
    case "equals":
      return scored.filter((s) => same(s.v, x)).map((s) => s.item);
    case "notEquals":
      return scored.filter((s) => !same(s.v, x)).map((s) => s.item);
    case "greaterThan":
      return scored.filter((s) => s.v > x).map((s) => s.item);
    case "greaterOrEqual":
      return scored.filter((s) => s.v >= x).map((s) => s.item);
    case "lessThan":
      return scored.filter((s) => s.v < x).map((s) => s.item);
    case "lessOrEqual":
      return scored.filter((s) => s.v <= x).map((s) => s.item);
    case "between":
      return scored.filter((s) => s.v >= x && s.v <= y).map((s) => s.item);
    case "notBetween":
      return scored.filter((s) => s.v < x || s.v > y).map((s) => s.item);
    case "top":
    case "bottom": {
      const ranked = scored.slice().sort((a, b) => (filter.operator === "top" ? b.v - a.v : a.v - b.v));
      const by = filter.topBy ?? "items";
      if (by === "items") return ranked.slice(0, Math.max(0, Math.floor(x))).map((s) => s.item);
      // Percent: the fewest items whose values make up that share of the total.
      // Sum: the fewest items whose values add up to it.
      const total = ranked.reduce((sum, s) => sum + s.v, 0);
      const target = by === "percent" ? (total * x) / 100 : x;
      const out: T[] = [];
      let running = 0;
      for (const s of ranked) {
        if (out.length > 0 && running >= target) break;
        out.push(s.item);
        running += s.v;
      }
      return out;
    }
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
  const axisSet = (fields: PlacedField[], options: LayoutOptions, grand: boolean): string => {
    const build = (k: number): string => {
      const rest = fields.slice(k);
      const inner = k + 1 < fields.length ? `CrossJoin(${detail(fields[k])}, ${build(k + 1)})` : detail(fields[k]);
      const wanted = k === 0 ? grand : options.subtotals;
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
      ? `CrossJoin(${axisSet(layout.columns, layout.options, layout.options.grandTotalColumn ?? layout.options.grandTotals)}, ${measures})`
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
    lines.push(`  ${nonEmpty}${axisSet(layout.rows, layout.options, layout.options.grandTotalRow ?? layout.options.grandTotals)} ON ROWS`);
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

    async details(connectionId, request) {
      await wait(300);
      need(connectionId);
      needCube(request.catalog, request.cube);
      return cube.details(request);
    },

    async cancel(connectionId) {
      const set = running.get(connectionId);
      if (set) {
        for (const stop of Array.from(set)) stop(fail("cancelled", 499, "The query was cancelled."));
      }
    },
  };
}
