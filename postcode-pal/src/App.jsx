import { useMemo, useState } from "react";
import {
  Alert,
  AppBar,
  Box,
  Button,
  Card,
  CardContent,
  Chip,
  Container,
  CssBaseline,
  Paper,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TextField,
  ThemeProvider,
  Toolbar,
  Tooltip,
  Typography,
  createTheme,
} from "@mui/material";
import FileDownloadIcon from "@mui/icons-material/FileDownload";
import ArrowDownwardIcon from "@mui/icons-material/ArrowDownward";
import ArrowUpwardIcon from "@mui/icons-material/ArrowUpward";
import MarkunreadMailboxIcon from "@mui/icons-material/MarkunreadMailbox";
import * as XLSX from "xlsx";

const API_BASE_URL =
  "https://list.melissadata.net/v1/Business/rest/Service.svc/get/json/zip";

const theme = createTheme({
  palette: {
    mode: "light",
    primary: { main: "#2596be" },
    secondary: { main: "#67BED9" },
    background: { default: "#F4F2ED", paper: "#FFFFFF" },
    success: { main: "#3F6B63" },
  },
  typography: {
    fontFamily: '"Inter", "Helvetica", "Arial", sans-serif',
    h4: { fontWeight: 700, letterSpacing: "-0.01em" },
    h6: { fontWeight: 600 },
  },
  shape: { borderRadius: 10 },
  components: {
    MuiTableCell: {
      styleOverrides: {
        head: { fontWeight: 700, backgroundColor: "#EEF1F0" },
      },
    },
  },
});

const codeSx = {
  fontFamily: '"Roboto Mono", "Menlo", monospace',
  fontVariantNumeric: "tabular-nums",
};

export default function ZipBusinessLookup() {
  const [apiKey, setApiKey] = useState("");
  const [zip, setZip] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [sortDir, setSortDir] = useState("desc");

  const rows = useMemo(() => {
    if (!result) return [];
    const sorted = [...result.rows].sort((a, b) =>
      sortDir === "desc" ? b.count - a.count : a.count - b.count,
    );
    return sorted;
  }, [result, sortDir]);

  const canSearch = apiKey.trim().length > 0 && /^\d{5}$/.test(zip.trim());

  async function handleSearch(e) {
    e?.preventDefault();
    if (!canSearch || loading) return;
    setLoading(true);
    setError("");
    setResult(null);

    try {
      const url = `${API_BASE_URL}?id=${encodeURIComponent(
        apiKey.trim(),
      )}&zip=${encodeURIComponent(zip.trim())}`;
      const res = await fetch(url);
      if (!res.ok) {
        throw new Error(`Request failed with status ${res.status}`);
      }
      const data = await res.json();
      const business = data?.Business;
      const statusCode = business?.Result?.StatusCode ?? "Unknown";
      const totalCount = Number(business?.TotalCount?.Count ?? 0);
      const streetList = business?.CountDetails?.StreetRange?.Street ?? [];
      const list = Array.isArray(streetList) ? streetList : [streetList];

      const parsedRows = list.filter(Boolean).map((item) => ({
        geography: item.Geography,
        count: Number(item.Count ?? 0),
      }));

      if (statusCode !== "Approved") {
        setError(
          `API returned status "${statusCode}" for ZIP ${zip.trim()}. Double-check the ZIP code and your License Key.`,
        );
      }

      setResult({
        zip: zip.trim(),
        totalCount,
        rows: parsedRows,
        statusCode,
      });
    } catch (err) {
      setError(
        err instanceof TypeError
          ? "Network request was blocked (likely CORS). Route this through a backend proxy — see the setup note above."
          : err.message || "Something went wrong looking up that ZIP code.",
      );
    } finally {
      setLoading(false);
    }
  }

  function handlePrint() {
    window.print();
  }

  function handleExportExcel() {
    if (!result) return;

    const sheetData = rows.map((r) => ({
      "ZIPCODE+4": r.geography,
      "Business Count": r.count,
    }));
    sheetData.push({
      "ZIPCODE+4": "Total",
      "Business Count": result.totalCount,
    });

    const worksheet = XLSX.utils.json_to_sheet(sheetData);
    worksheet["!cols"] = [{ wch: 14 }, { wch: 16 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `ZIP ${result.zip}`);

    XLSX.writeFile(workbook, `zip-${result.zip}-businesses.xlsx`);
  }

  function toCsv(rows, totalCount) {
    const header = ["ZIPCODE", "Business Count"];
    const lines = [header.join(",")];
    rows.forEach((r) => lines.push(`${r.geography},${r.count}`));
    lines.push("");
    lines.push(`Total,${totalCount}`);
    return lines.join("\n");
  }

  function downloadCsv(filename, csvText) {
    const blob = new Blob([csvText], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <AppBar
        position="static"
        elevation={0}
        color="primary"
        className="no-print"
      >
        <Toolbar>
          <MarkunreadMailboxIcon sx={{ mr: 1.5 }} />
          <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
            Melissa ZIP Business Lookup
          </Typography>
        </Toolbar>
      </AppBar>

      <Box
        sx={{
          bgcolor: "background.default",
          minHeight: "100vh",
          py: 5,
        }}
      >
        <Container maxWidth="md">
          <Typography variant="h4" gutterBottom className="no-print">
            Look up businesses by ZIP code
          </Typography>
          <Typography
            variant="body1"
            color="text.secondary"
            sx={{ mb: 4 }}
            className="no-print"
          >
            Enter a 5-digit ZIP code to see the total number of registered
            businesses and how they're distributed across ZIP Codes.
          </Typography>

          <Paper variant="outlined" sx={{ p: 3, mb: 4 }} className="no-print">
            <Box component="form" onSubmit={handleSearch}>
              <Stack
                direction={{ xs: "column", sm: "row" }}
                spacing={2}
                alignItems={{ sm: "flex-start" }}
              >
                <TextField
                  label="Melissa License Key"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  showKey={true}
                  fullWidth
                  size="small"
                />
                <TextField
                  label="ZIP code"
                  value={zip}
                  onChange={(e) =>
                    setZip(e.target.value.replace(/[^\d]/g, "").slice(0, 5))
                  }
                  placeholder="30090"
                  size="small"
                  sx={{ minWidth: { sm: 160 } }}
                  inputProps={{ inputMode: "numeric", maxLength: 5 }}
                />
                <Button
                  type="submit"
                  variant="contained"
                  disabled={!canSearch || loading}
                  sx={{ height: "2.5rem", whiteSpace: "nowrap" }}
                >
                  {loading ? "Searching…" : "Search"}
                </Button>
              </Stack>
            </Box>
          </Paper>

          {error && (
            <Alert severity="warning" sx={{ mb: 3 }}>
              {error}
            </Alert>
          )}

          {result && (
            <div>
              <Card variant="outlined" sx={{ mb: 3 }}>
                <CardContent>
                  <Stack
                    direction={{ xs: "column", sm: "row" }}
                    justifyContent="space-between"
                    alignItems={{ sm: "center" }}
                    spacing={1}
                  >
                    <Box>
                      <Typography variant="overline" color="text.secondary">
                        ZIP {result.zip}
                      </Typography>
                      <Typography variant="h3" sx={codeSx}>
                        {result.totalCount.toLocaleString()}
                      </Typography>
                      <Typography variant="body2" color="text.secondary">
                        total businesses
                      </Typography>
                    </Box>
                    <Chip
                      label={result.statusCode}
                      color={
                        result.statusCode === "Approved" ? "success" : "warning"
                      }
                      variant="outlined"
                    />
                  </Stack>
                </CardContent>
              </Card>

              <Stack
                direction="row"
                justifyContent="space-between"
                alignItems="center"
                sx={{ mb: 1 }}
                className="no-print"
              >
                <Typography variant="h6">ZIPCODE breakdown</Typography>
                <span>
                  <Tooltip title="Print">
                    <Button
                      size="small"
                      startIcon={<FileDownloadIcon />}
                      onClick={handlePrint}
                      disabled={rows.length === 0}
                    >
                      Print
                    </Button>
                  </Tooltip>
                  <Tooltip title="Export to Excel">
                    <Button
                      size="small"
                      onClick={handleExportExcel}
                      disabled={rows.length === 0}
                      startIcon={<FileDownloadIcon />}
                      className="no-print"
                    >
                      Export Excel
                    </Button>
                  </Tooltip>
                  <Tooltip title="Export to CSV">
                    <Button
                      size="small"
                      onClick={() =>
                        downloadCsv(
                          `zip-${result.zip}-businesses.csv`,
                          toCsv(rows, result.totalCount),
                        )
                      }
                      disabled={rows.length === 0}
                      startIcon={<FileDownloadIcon />}
                      className="no-print"
                    >
                      Export CSV
                    </Button>
                  </Tooltip>
                </span>
              </Stack>

              <TableContainer component={Paper} variant="outlined">
                <Table size="small">
                  <TableHead>
                    <TableRow>
                      <TableCell sx={codeSx}>ZIPCODE+4</TableCell>
                      <TableCell
                        align="right"
                        onClick={() =>
                          setSortDir((d) => (d === "desc" ? "asc" : "desc"))
                        }
                        sx={{ cursor: "pointer", userSelect: "none" }}
                      >
                        <Stack
                          direction="row"
                          spacing={0.5}
                          display="flex"
                          justifyContent="center"
                          alignItems="center"
                        >
                          <span>Business count</span>
                          {sortDir === "desc" ? (
                            <ArrowDownwardIcon
                              sx={{ fontSize: "1.25rem", pt: "0.25rem" }}
                            />
                          ) : (
                            <ArrowUpwardIcon
                              sx={{ fontSize: "1.25rem", pb: "0.25rem" }}
                            />
                          )}
                        </Stack>
                      </TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {rows.map((r) => (
                      <TableRow key={r.geography} hover>
                        <TableCell sx={codeSx}>{r.geography}</TableCell>
                        <TableCell align="right">{r.count}</TableCell>
                      </TableRow>
                    ))}
                    {rows.length === 0 && (
                      <TableRow>
                        <TableCell colSpan={2} align="center">
                          <Typography color="text.secondary" sx={{ py: 2 }}>
                            No ZIPCODE+4 records were returned for this ZIP
                            code.
                          </Typography>
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </TableContainer>
            </div>
          )}
        </Container>
      </Box>
    </ThemeProvider>
  );
}
