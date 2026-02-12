# excelize-simple-builder

This package provides a simple way to generate Excel reports from Go struct slices.
It supports flat column lists and nested column groups with automatic header cell merging.

## Features

- Flat column lists
- Nested column groups  with automatic header cell merging
- Custom styles for table header and body
- Type-safe data access through getter functions

## Usage Examples

### Flat Column List

```go
type Product struct {
    ID    string
    Name  string
    Price float64
    Stock int
}

products := []Product{
    {ID: "P001", Name: "Laptop", Price: 999.99, Stock: 15},
    {ID: "P002", Name: "Mouse", Price: 29.99, Stock: 50},
}

cols := []esb.Column{
    esb.Col("ID", func(p Product) string { return p.ID }),
    esb.Col("Name", func(p Product) string { return p.Name }),
    esb.Col("Price", func(p Product) float64 { return p.Price }),
    esb.Col("Stock", func(p Product) int { return p.Stock }),
}

f := excelize.NewFile()
defer f.Close()

builder := esb.New[Product](cols)
f, _ = builder.WriteTo(f, "Products", products)
_ = f.SaveAs("products.xlsx")
```

Result in Excel:

<table>
    <tr>
        <th>ID</th>
        <th>Name</th>
        <th>Price</th>
        <th>Stock</th>
    </tr>
    <tr>
        <td>P001</td>
        <td>Laptop</td>
        <td>999.99</td>
        <td>15</td>
    </tr>
    <tr>
        <td>P002</td>
        <td>Mouse</td>
        <td>29.99</td>
        <td>50</td>
    </tr>
</table>

### Nested Groups

```go
type Server struct {
    ID     string
    Name   string
    Config Config
    Status string
}

type Config struct {
    CPU    CPU
    Memory Memory
}

type CPU struct {
    Count uint32
    Cores uint32
}

type Memory struct {
    Size uint32
    Type string
}

servers := []Server{
    {
        ID: "srv-001",
        Name: "Web Server",
        Config: Config{
            CPU:    CPU{Count: 2, Cores: 8},
            Memory: Memory{Size: 32, Type: "DDR4"},
        },
        Status: "Running",
    },
}

cols := []esb.Column{
    esb.Col("ID", func(s Server) string { return s.ID }),
    esb.Col("Name", func(s Server) string { return s.Name }),
    esb.Group("Configuration", func(s Server) Config { return s.Config },
        esb.Group("CPU", func(c Config) CPU { return c.CPU },
            esb.Col("Count", func(cpu CPU) uint32 { return cpu.Count }),
            esb.Col("Cores", func(cpu CPU) uint32 { return cpu.Cores }),
        ),
        esb.Group("Memory", func(c Config) Memory { return c.Memory },
            esb.Col("Size", func(m Memory) uint32 { return m.Size }),
            esb.Col("Type", func(m Memory) string { return m.Type }),
        ),
    ),
    esb.Col("Status", func(s Server) string { return s.Status }),
}

f := excelize.NewFile()
defer f.Close()

builder := esb.New[Server](cols)
f, _ = builder.WriteTo(f, "Servers", servers)
_ = f.SaveAs("servers.xlsx")
```

Result in Excel:

<table>
    <tr>
        <th rowspan="3">ID</th>
        <th rowspan="3">Name</th>
        <th colspan="4">Configuration</th>
        <th rowspan="3">Status</th>
    </tr>
    <tr>
        <th colspan="2">CPU</th>
        <th colspan="2">Memory</th>
    </tr>
    <tr>
        <th>Count</th>
        <th>Cores</th>
        <th>Size GB</th>
        <th>Type</th>
    </tr>
    <tr>
        <td rowspan="3">srv-001</td>
        <td rowspan="3">Web Server</td>
        <td>2</td>
        <td>8</td>
        <td>32</td>
        <td>DDR4</td>
        <td rowspan="3">Running</td>
    </tr>
</table>

### Custom Styles

```go
r := esb.New(cols,
    esb.WithHeaderStyle[Product](func(s *excelize.Style) {
        s.Font = &excelize.Font{Size: 14, Bold: true}
        s.Border = []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
        }
    }),
    esb.WithBodyStyle[Product](func(s *excelize.Style) {
        s.Font = &excelize.Font{Size: 12}
    }),
)
```