# Supplying subreports

The converter never resolves an RDL `ReportName` through SSRS, the filesystem, a database, or the network.
A parent report that contains `Subreport` items can render only when the same render request contains:

1. the exact child RDL bytes;
2. one concrete data instance for every child-parameter combination the parent can invoke; and
3. every dataset row that child instance needs.

This contract is identical for `POST /v1/render` and `createConverter().render()`. Bundled subreports are
supported for `PDF`, `DOCX_EDITABLE`, and `DOCX_VISUAL`. `XLSX` fails closed.

## Recommended preparation workflow

1. Analyze the parent RDL. Its `subreports` array identifies each `ReportName`, item name, and parameter
   expression. A parent with an unresolved subreport reports `compatible: false`; that is expected because
   analysis has no render-time bundle.
2. Obtain each referenced child RDL from the caller's trusted report catalogue.
3. Analyze every child RDL separately to obtain its declared parameters, rendering datasets, and exact
   `DataField` names.
4. Execute data access outside this service. Group the resulting child rows by the values of the child's
   declared report parameters.
5. Put every child definition and invocation instance into the parent render request.

RDL query text remains metadata. Neither analyzing nor rendering a bundle executes it.

For example, parent analysis exposes the call site:

```json
{
  "compatible": false,
  "subreports": [
    {
      "name": "RiskObjectives",
      "reportName": "/Risk/Objective Subreport",
      "parameters": [
        {
          "name": "EntityID",
          "value": "=Fields!PRiskInstID.Value"
        }
      ]
    }
  ]
}
```

Analyze `Objective Subreport.rdl` separately. Its analysis tells the caller which child parameters and
dataset fields to populate. Do not infer child fields from the parent analysis.

## Request shape

```ts
type SubreportBundle = Record<string, {
  // Complete child .rdl bytes encoded as standard base64.
  rdlBase64: string;

  // Exactly one entry per unique combination of the child's declared report parameters.
  instances: Array<{
    parameters: Record<string, unknown>;
    datasets: Record<string, Array<Record<string, unknown>>>;
  }>;
}>;
```

The object key is the child's RDL `ReportName`. Matching is case-insensitive, normalizes `/` and `\`,
collapses repeated slashes, and ignores a trailing `.rdl`. Use the RDL value unchanged when practical:

```xml
<Subreport Name="RiskObjectives">
  <ReportName>/Risk/Objective Subreport</ReportName>
  <Parameters>
    <Parameter Name="EntityID">
      <Value>=Fields!PRiskInstID.Value</Value>
    </Parameter>
  </Parameters>
</Subreport>
```

Its request definition is keyed by `/Risk/Objective Subreport`.

The mapping between the two RDLs and the request is:

| RDL declaration | Request location |
| --- | --- |
| Parent `Subreport/ReportName` | Key under top-level `subreports` |
| Child RDL bytes | `subreports[reportName].rdlBase64` |
| Parent `Subreport/Parameters/Parameter@Name` | `instances[].parameters` key |
| Evaluated parent parameter value | `instances[].parameters[name]` value |
| Child `DataSet@Name` | `instances[].datasets` key |
| Child `Field/DataField` | Property name in each child dataset row |

## Complete HTTP JSON example

Assume the parent dataset invokes the child twice, for entity IDs `42` and `84`:

```json
{
  "rdlBase64": "BASE64_PARENT_RDL",
  "output": "PDF",
  "outputFileName": "risk-register",
  "parameters": {
    "ReportYear": 2026
  },
  "datasets": {
    "MainData": [
      { "Risk Instance ID": 42, "Risk Name": "Supplier concentration" },
      { "Risk Instance ID": 84, "Risk Name": "Water interruption" }
    ]
  },
  "subreports": {
    "/Risk/Objective Subreport": {
      "rdlBase64": "BASE64_CHILD_RDL",
      "instances": [
        {
          "parameters": {
            "EntityID": 42
          },
          "datasets": {
            "ObjectiveData": [
              {
                "Objective ID": 1001,
                "Objective Name": "Maintain supply resilience",
                "From Risk Instance ID": 42
              }
            ]
          }
        },
        {
          "parameters": {
            "EntityID": 84
          },
          "datasets": {
            "ObjectiveData": []
          }
        }
      ]
    }
  }
}
```

Important details:

- `MainData`, `ObjectiveData`, and every other dataset key are exact RDL `DataSet Name` values.
- Row properties such as `Risk Instance ID` and `Objective Name` are exact `DataField` values, including
  case, spaces, and punctuation. They are not the internal RDL `Field Name` used in expressions.
- An invocation with no rows still needs its instance and an empty dataset array, as shown for entity `84`.
- Do not create two instances with the same declared child-parameter values. Duplicate signatures are
  rejected even if their row arrays differ.

Send the JSON normally:

```bash
curl -X POST http://localhost:7070/v1/render \
  -H 'content-type: application/json' \
  --data @request.json \
  --output risk-register.pdf
```

## Multipart HTTP example

For multipart, the parent RDL remains the `rdl` file part. The `request` JSON part contains `subreports`,
including each child RDL as base64:

```bash
curl -X POST http://localhost:7070/v1/render \
  -F 'rdl=@parent.rdl;type=application/xml' \
  -F 'request=@render-request.json;type=application/json' \
  --output risk-register.pdf
```

`render-request.json` has the same shape as the JSON example except it omits the top-level `rdlBase64`.

## Library example

```js
import fs from 'node:fs/promises';
import { createConverter } from 'rdl-converter-service';

const [parentRdl, childRdl] = await Promise.all([
  fs.readFile('parent.rdl'),
  fs.readFile('objective-subreport.rdl'),
]);

const converter = await createConverter();
try {
  const rendered = await converter.render({
    rdl: parentRdl,
    output: 'PDF',
    parameters: { ReportYear: 2026 },
    datasets: {
      MainData: [
        { 'Risk Instance ID': 42, 'Risk Name': 'Supplier concentration' },
      ],
    },
    subreports: {
      '/Risk/Objective Subreport': {
        rdlBase64: childRdl.toString('base64'),
        instances: [
          {
            parameters: { EntityID: 42 },
            datasets: {
              ObjectiveData: [
                {
                  'Objective ID': 1001,
                  'Objective Name': 'Maintain supply resilience',
                  'From Risk Instance ID': 42,
                },
              ],
            },
          },
        ],
      },
    },
  });

  await fs.writeFile('risk-register.pdf', rendered.buffer);
} finally {
  await converter.close();
}
```

The only transport difference is the parent RDL: the library accepts a `Buffer`, `Uint8Array`, or XML
string as `rdl`; the HTTP JSON endpoint accepts it as `rdlBase64`. Child RDLs use `rdlBase64` in both APIs.

## How invocation matching works

For every visible parent subreport item, the converter:

1. evaluates the parent `<Parameter><Value>` expressions in the current parent row/group scope;
2. resolves the child's declared defaults and validates its parameter types;
3. creates a canonical signature from the child's declared parameters; and
4. looks up exactly one supplied instance with that signature.

Canonical matching follows the child parameter type:

| Child parameter type | Matching form |
| --- | --- |
| `String` | String value |
| `Integer`, `Float` | Numeric value (`"42"` and `42` match an integer parameter) |
| `Boolean` | Boolean value (`"true"` and `true` match) |
| `DateTime` | ISO timestamp |
| Multi-value | Ordered array of canonical element values |

Only parameters declared by the child participate in the signature. Use the exact child parameter names.
If a child has no report parameters, it has one possible signature and therefore accepts exactly one
instance.

When the same parameter signature is referenced by several parent rows, supply one instance; it is safely
reused. If two calls need different data, their declared child parameter values must differ.

## Preparing instances from query results

Data access belongs to the caller. A typical caller groups already-authorized child rows by the parameter
the child RDL declares:

```js
function instancesByEntity(rows) {
  const grouped = new Map();
  for (const row of rows) {
    const entityID = Number(row['From Risk Instance ID']);
    if (!grouped.has(entityID)) grouped.set(entityID, []);
    grouped.get(entityID).push(row);
  }
  return [...grouped].map(([EntityID, childRows]) => ({
    parameters: { EntityID },
    datasets: { ObjectiveData: childRows },
  }));
}
```

Also create instances with empty row arrays for parent parameter values that have no child result. The
renderer must be able to distinguish “the child query returned no rows” from “the caller forgot to supply
this invocation.”

## Nested child and grandchild reports

All definitions live in the same top-level `subreports` map. If `/Reports/Child` contains a subreport whose
`ReportName` is `/Reports/Grandchild`, supply both:

```json
{
  "subreports": {
    "/Reports/Child": {
      "rdlBase64": "BASE64_CHILD_RDL",
      "instances": [
        {
          "parameters": { "ParentID": 7 },
          "datasets": { "ChildData": [{ "Parent ID": 7, "Child ID": 70 }] }
        }
      ]
    },
    "/Reports/Grandchild": {
      "rdlBase64": "BASE64_GRANDCHILD_RDL",
      "instances": [
        {
          "parameters": { "ChildID": 70 },
          "datasets": { "GrandchildData": [{ "Child ID": 70, "Evidence": "Complete" }] }
        }
      ]
    }
  }
}
```

The grandchild's parameter expression is evaluated in the child row scope. Definitions are resolved
recursively, while invocation data is still selected only by the declared parameters of the report being
called.

## Validation and limits

The resolver fails closed:

| Condition | Error |
| --- | --- |
| Referenced `ReportName` has no bundle definition | `UNSUPPORTED_FEATURE` |
| PDF/visual-DOCX invocation has no matching instance | `DATASET_MISSING` |
| Required child dataset is absent | `DATASET_MISSING` |
| A child row omits an exact `DataField` | `FIELD_MISSING` |
| Child parameter is missing or has the wrong type | `PARAMETER_INVALID` |
| Duplicate invocation signature, unused definition, malformed base64, or cycle | `RDL_INVALID` |
| Unsupported child construct/output mode or nesting beyond eight levels | `UNSUPPORTED_FEATURE` |

Additional constraints:

- at most 32 bundled report definitions;
- each child RDL obeys `RDL_MAX_RDL_BYTES`;
- combined child RDL bytes obey the complete request-size limit;
- parent rows plus rows from every supplied child instance obey `RDL_MAX_ROWS`;
- child fonts join the parent font requirement and remain fail-closed in strict font mode; and
- the current child body may contain only tablix report items.

Every supplied definition must be referenced by the parent/child RDL graph. Extra definitions are rejected
instead of being silently ignored.
