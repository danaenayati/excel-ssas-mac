// Database and cube selects for the report being edited.

import type { Loadable } from "../state/useCached";
import type { Catalog, Cube } from "../types";
import { ErrorNote } from "./common";

interface Props {
  catalogs: Loadable<Catalog[]>;
  cubes: Loadable<Cube[]>;
  catalog: string;
  cube: string;
  disabled?: boolean;
  onCatalog: (name: string) => void;
  onCube: (name: string) => void;
}

export function CubePicker({ catalogs, cubes, catalog, cube, disabled, onCatalog, onCube }: Props) {
  const catalogList = catalogs.data ?? [];
  const cubeList = cubes.data ?? [];
  // A report made elsewhere may name a database or cube this connection does not have.
  const catalogMissing = catalog !== "" && catalogs.data !== null && !catalogList.some((c) => c.name === catalog);
  const cubeMissing = cube !== "" && cubes.data !== null && !cubeList.some((c) => c.name === cube);

  return (
    <div className="cube-picker">
      <label className="row-field">
        <span className="row-label">Database</span>
        <select
          className="select"
          value={catalog}
          disabled={disabled || catalogs.loading || catalogs.data === null}
          onChange={(e) => onCatalog(e.target.value)}
        >
          {catalogs.loading && <option value={catalog}>Loading…</option>}
          {!catalogs.loading && catalog === "" && <option value="">Choose a database</option>}
          {catalogMissing && <option value={catalog}>{catalog} (not found)</option>}
          {catalogList.map((c) => (
            <option key={c.name} value={c.name} title={c.description || undefined}>
              {c.name}
            </option>
          ))}
        </select>
      </label>
      <label className="row-field">
        <span className="row-label">Cube</span>
        <select
          className="select"
          value={cube}
          disabled={disabled || catalog === "" || cubes.loading || cubes.data === null}
          onChange={(e) => onCube(e.target.value)}
        >
          {cubes.loading && <option value={cube}>Loading…</option>}
          {!cubes.loading && cube === "" && <option value="">Choose a cube</option>}
          {cubeMissing && <option value={cube}>{cube} (not found)</option>}
          {cubeList.map((c) => (
            <option key={c.name} value={c.name}>
              {c.caption || c.name}
            </option>
          ))}
        </select>
      </label>
      {catalogs.error && (
        <ErrorNote error={catalogs.error} action={{ label: "Try again", run: catalogs.retry }} />
      )}
      {!catalogs.error && cubes.error && (
        <ErrorNote error={cubes.error} action={{ label: "Try again", run: cubes.retry }} />
      )}
    </div>
  );
}
