# Tableau Metadata Extractor

A Python tool to extract and document metadata from Tableau `.twb` and `.twbx` workbook files. It parses the XML structure, identifies calculated fields, tracks lineage, and exports everything to a user-friendly Excel file.

---

## Features

- Extracts metadata from `.twb` and `.twbx` files
- Lists datasources, worksheets, dashboards, parameters, and fields
- Lists fields in rows, columns, and filter shelfs in each worksheet 
- Identifies calculated fields and their formulas
- Tracks lineage: which fields are referenced in each calculation
- Exports all metadata to an Excel workbook with separate sheets

---

## Requirements

Install dependencies using:

```bash
pip install -r requirements.txt

---

## Excel Output Example

```markdown
### 🗂 Worksheets
| Worksheet Name     |
|--------------------|
| Sales Overview     |
| Profit Trends      |
| Regional Summary   |

### Dashboards
| Dashboard Name     |
|--------------------|
| Executive Overview |
| Regional Insights  |

### Parameters
| Name           | Data Type | Current Value |
|----------------|-----------|----------------|
| Region Filter  | String    | West           |
| Date Range     | Date      | 2023-01-01     |

### Fields
| Name           | Caption       | Data Type | Role     | Type        | Formula                        | Is Calculated | Data Source |
|----------------|---------------|-----------|----------|-------------|--------------------------------|----------------|-------------|
| [Sales]        | Sales         | Float     | Measure  | Quantitative|                                | False          | Orders      |
| [Profit Ratio] | Profit Ratio  | Float     | Measure  | Calculated  | SUM([Profit])/SUM([Sales])     | True           | Orders      |
| [Year]         | Year          | Integer   | Dimension| Calculated  | DATEPART('year', [Order Date]) | True           | Orders      |

### Calculated Fields
| Field Name     | Caption       | Formula                          | Data Source | References                     |
|----------------|---------------|----------------------------------|-------------|--------------------------------|
| [Profit Ratio] | Profit Ratio  | SUM([Profit])/SUM([Sales])       | Orders      | [Profit], [Sales]              |
| [Year]         | Year          | DATEPART('year', [Order Date])   | Orders      | [Order Date]                   |

### Lineage
| Calculated Field | Depends On     | Data Source |
|------------------|----------------|-------------|
| [Profit Ratio]   | [Profit]       | Orders      |
| [Profit Ratio]   | [Sales]        | Orders      |
| [Year]           | [Order Date]   | Orders      |
