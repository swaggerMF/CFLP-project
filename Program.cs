using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ClosedXML.Excel;

class Course
{
    public string Name { get; set; }
    public int? Year { get; set; }          // null dacă lipsește
    public int? Semester { get; set; }      // null dacă lipsește
    public double? Credits { get; set; }    // null dacă lipsește (de regulă nu ar trebui pt UPT)
    public string Grade { get; set; }  // nota/grade din coloana 9

    public Course()
    {
        Name = string.Empty;
        Grade = string.Empty;
    }

    public Course(string name, int? year, int? semester, double? credits, string grade = "")
    {
        if (name == null)
        {
            Name = string.Empty;
        }
        else
        {
            Name = name;
        }
        Year = year;
        Semester = semester;
        Credits = credits;
        if (grade == null)
        {
            Grade = string.Empty;
        }
        else
        {
            Grade = grade;
        }
    }
}

class Group
{
    public int Id { get; set; }                 // 1..n (Nr. crt pt PV)
    public string HostName { get; set; }
    public double HostCredits { get; set; }     // creditul gazdei din primul rând al grupului
    public List<Course> UptCourses { get; set; }

    public Group()
    {
        HostName = string.Empty;
        UptCourses = new List<Course>();
    }

    public Group(int id, string hostName, double hostCredits, List<Course>? uptCourses = null)
    {
        Id = id;
        if (hostName == null)
        {
            HostName = string.Empty;
        }
        else
        {
            HostName = hostName;
        }
        HostCredits = hostCredits;
        if (uptCourses != null)
        {
            UptCourses = uptCourses;
        }
        else
        {
            UptCourses = new List<Course>();
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        string excelPath = string.Empty;
        if (args.Length > 0)
        {
            excelPath = args[0];
        }
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            excelPath = AutoFindExcel();
            if (string.IsNullOrWhiteSpace(excelPath) == false)
            {
                Console.WriteLine("Gasit automat fisier Excel: " + excelPath);
            }
        }

        if (string.IsNullOrWhiteSpace(excelPath))
        {
            Console.Write("Cale Excel (.xlsx): ");
            var input = Console.ReadLine();
            if (input != null)
            {
                excelPath = input;
            }
        }

        excelPath = excelPath.Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(excelPath) || File.Exists(excelPath) == false)
        {
            Console.WriteLine("Fisierul Excel nu a putut fi gasit. Oprire.");
            return;
        }

        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet(1);

        var groups = BuildGroups(ws);
        Console.WriteLine($"Incarcat {groups.Count} grupuri din {excelPath}.");
        PrintHelp();
        CommandLoop(groups);
    }

    static List<Group> BuildGroups(IXLWorksheet ws)
    {
        // Coloane (1-based)
        const int COL_NR_CRT = 1;
        const int COL_UPT_NAME = 2;
        const int COL_UPT_YEAR = 3;
        const int COL_UPT_SEM = 4;
        const int COL_UPT_CREDITS = 5;
        const int COL_HOST_NAME = 6;
        const int COL_HOST_CREDITS = 7;
        const int COL_NOTE = 9; // coloana cu grade

        var groups = new List<Group>();
        Group? current = null;
        int nextId = 1;

        foreach (var row in ws.RowsUsed())
        {
            string uptName = row.Cell(COL_UPT_NAME).GetString().Trim();
            string hostName = row.Cell(COL_HOST_NAME).GetString().Trim();

            // Stop la TOTAL (dacă există)
            if (string.IsNullOrWhiteSpace(uptName) == false &&
                uptName.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                break;

            // Ignoră rând complet gol (în coloanele relevante)
            bool allEmpty =
                string.IsNullOrWhiteSpace(row.Cell(COL_NR_CRT).GetString()) &&
                string.IsNullOrWhiteSpace(uptName) &&
                string.IsNullOrWhiteSpace(row.Cell(COL_UPT_CREDITS).GetString()) &&
                string.IsNullOrWhiteSpace(hostName) &&
                string.IsNullOrWhiteSpace(row.Cell(COL_HOST_CREDITS).GetString()) &&
                string.IsNullOrWhiteSpace(row.Cell(COL_NOTE).GetString());
            if (allEmpty) continue;

            double? hostCredits = ReadIntOrNull(row.Cell(COL_HOST_CREDITS));
            double? uptCredits  = ReadIntOrNull(row.Cell(COL_UPT_CREDITS));
            int? year = ReadIntOrNull(row.Cell(COL_UPT_YEAR));
            int? sem  = ReadIntOrNull(row.Cell(COL_UPT_SEM));
            string grade = row.Cell(COL_NOTE).GetString().Trim();

            if (hostCredits.HasValue)
            {
                var resolvedHost = hostName;
                if (string.IsNullOrWhiteSpace(resolvedHost))
                {
                    resolvedHost = "(HOST NECUNOSCUT)";
                }
                current = new Group(
                    nextId++,
                    resolvedHost,
                    hostCredits.Value);
                groups.Add(current);
            }
            else
            {
                if (current == null)
                    continue;
                if (string.IsNullOrWhiteSpace(hostName) == false &&
                    (string.IsNullOrWhiteSpace(current.HostName) || current.HostName == "(HOST NECUNOSCUT)"))
                {
                    current.HostName = hostName;
                }
            }
            if (current != null && string.IsNullOrWhiteSpace(uptName) == false && uptCredits.HasValue)
            {
                current.UptCourses.Add(new Course(uptName, year, sem, uptCredits, grade));
            }
        }

        return groups;
    }

    static void CommandLoop(List<Group> groups)
    {
        while (true)
        {
            Console.Write("cmd> ");
            var lineInput = Console.ReadLine();
            var line = string.Empty;
            if (lineInput != null)
            {
                line = lineInput;
            }
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = line.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var cmd = parts[0].ToLowerInvariant();

            switch (cmd)
            {
                case "list":
                    PrintGroups(groups);
                    break;
                case "check":
                    CheckGroups(groups);
                    break;
                case "set-host-name":
                    SetHostName(groups, parts);
                    break;
                case "set-host-credits":
                    SetHostCredits(groups, parts);
                    break;
                case "set-upt-name":
                    SetUptName(groups, parts);
                    break;
                case "set-upt-credits":
                    SetUptCredits(groups, parts);
                    break;
                case "set-upt-grade":
                    SetUptGrade(groups, parts);
                    break;
                case "add-upt":
                    AddUpt(groups, parts);
                    break;
                case "remove-upt":
                    RemoveUpt(groups, parts);
                    break;
                case "split-upt":
                    SplitUpt(groups, parts);
                    break;
                case "add-group":
                    AddGroup(groups, parts);
                    break;
                case "export-excel":
                    ExportExcel(groups, parts);
                    break;
                case "export-docx":
                    ExportDocx(groups, parts);
                    break;
                case "help":
                    PrintHelp();
                    break;
                case "exit":
                    return;
                default:
                    Console.WriteLine("Comanda necunoscuta. tasteaza 'help' pentru lista.");
                    break;
            }
        }
    }

    static bool CheckGroups(List<Group> groups)
    {
        bool ok = true;
        foreach (var g in groups)
        {
            double uptSum = SumCredits(g.UptCourses);
            if (g.HostCredits < uptSum)
            {
                ok = false;
                Console.WriteLine($"Grup {g.Id}: INVALID (HostCredits={g.HostCredits} < suma UPT={uptSum}).");
            }
        }
        if (ok == true)
        {
            Console.WriteLine("Toate grupurile au HostCredits >= suma creditelor UPT.");
        }
        return ok;
    }

    static void PrintGroups(List<Group> groups)
    {
        foreach (var g in groups)
        {
            double uptSum = SumCredits(g.UptCourses);
            Console.WriteLine($"Group {g.Id}: HOST={g.HostName} (credite={g.HostCredits}) | UPT sum={uptSum} | {g.UptCourses.Count} discipline");
            for (int i = 0; i < g.UptCourses.Count; i++)
            {
                var c = g.UptCourses[i];
                Console.WriteLine($"  [{i + 1}] {c.Name} (An={c.Year}, Sem={c.Semester}, credite={c.Credits}, grade={c.Grade})");
            }
        }
    }

    static void SetHostName(List<Group> groups, string[] parts)
    {
        if (parts.Length < 3 || int.TryParse(parts[1], out var gid) == false)
        {
            Console.WriteLine("Folosire: set-host-name <groupId> <nume>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (group == null)
        {
            Console.WriteLine("GroupId inexistent.");
            return;
        }
        string name = string.Join(' ', parts.Skip(2));
        group.HostName = name;
        Console.WriteLine($"HostName grup {gid} setat la '{name}'.");
    }

    static void SetHostCredits(List<Group> groups, string[] parts)
    {
        if (parts.Length < 3 || int.TryParse(parts[1], out var gid) == false || double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out var credits) == false)
        {
            Console.WriteLine("Folosire: set-host-credits <groupId> <credite>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (group == null)
        {
            Console.WriteLine("GroupId inexistent.");
            return;
        }
        group.HostCredits = credits;
        Console.WriteLine($"HostCredits grup {gid} setat la {credits}.");
    }

    static void SetUptName(List<Group> groups, string[] parts)
    {
        if (parts.Length < 4 || int.TryParse(parts[1], out var gid) == false || int.TryParse(parts[2], out var idx) == false)
        {
            Console.WriteLine("Folosire: set-upt-name <groupId> <uptIndex> <nume>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (TryGetCourse(group, idx, out var course) == false)
        {
            Console.WriteLine("GroupId sau uptIndex invalid.");
            return;
        }
        string name = string.Join(' ', parts.Skip(3));
        course.Name = name;
        Console.WriteLine($"UPT[{idx}] din grup {gid} redenumita in '{name}'.");
    }

    static void SetUptCredits(List<Group> groups, string[] parts)
    {
        if (parts.Length < 4 || int.TryParse(parts[1], out var gid) == false || int.TryParse(parts[2], out var idx) == false ||
            double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out var credits) == false)
        {
            Console.WriteLine("Folosire: set-upt-credits <groupId> <uptIndex> <credite>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (TryGetCourse(group, idx, out var course) == false)
        {
            Console.WriteLine("GroupId sau uptIndex invalid.");
            return;
        }
        course.Credits = credits;
        Console.WriteLine($"Credite pentru UPT[{idx}] din grup {gid} setate la {credits}.");
    }

    static void SetUptGrade(List<Group> groups, string[] parts)
    {
        if (parts.Length < 4 || int.TryParse(parts[1], out var gid) == false || int.TryParse(parts[2], out var idx) == false)
        {
            Console.WriteLine("Folosire: set-upt-grade <groupId> <uptIndex> <grade>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (TryGetCourse(group, idx, out var course) == false)
        {
            Console.WriteLine("GroupId sau uptIndex invalid.");
            return;
        }
        string grade = string.Join(' ', parts.Skip(3));
        course.Grade = grade;
        Console.WriteLine($"Grade pentru UPT[{idx}] din grup {gid} setata la '{grade}'.");
    }

    static void AddUpt(List<Group> groups, string[] parts)
    {
        if (parts.Length < 4 || int.TryParse(parts[1], out var gid) == false)
        {
            Console.WriteLine("Folosire: add-upt <groupId> <nume> <credite> [an] [sem] [grade]");
            return;
        }
        if (double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out var credits) == false)
        {
            Console.WriteLine("Credite invalide.");
            return;
        }
        int? year = null;
        int? sem = null;
        if (parts.Length > 4 && int.TryParse(parts[4], out var y)) year = y;
        if (parts.Length > 5 && int.TryParse(parts[5], out var s)) sem = s;
        string grade = "";
        if (parts.Length > 6)
        {
            grade = parts[6];
        }

        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (group == null)
        {
            Console.WriteLine("GroupId inexistent.");
            return;
        }
        string name = parts[2];
        var course = new Course(name, year, sem, credits, grade);
        group.UptCourses.Add(course);
        Console.WriteLine($"Adaugat UPT '{name}' in grup {gid}.");
    }

    static void RemoveUpt(List<Group> groups, string[] parts)
    {
        if (parts.Length < 3 || int.TryParse(parts[1], out var gid) == false || int.TryParse(parts[2], out var idx) == false)
        {
            Console.WriteLine("Folosire: remove-upt <groupId> <uptIndex>");
            return;
        }
        var group = groups.FirstOrDefault(g => g.Id == gid);
        if (group == null || idx < 1 || idx > group.UptCourses.Count)
        {
            Console.WriteLine("GroupId sau uptIndex invalid.");
            return;
        }
        group.UptCourses.RemoveAt(idx - 1);
        Console.WriteLine($"UPT[{idx}] eliminata din grup {gid}.");
    }

    static void ExportExcel(List<Group> groups, string[] parts)
    {
        string fileName = "";
        if (parts.Length > 1)
        {
            fileName = parts[1];
        }
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = "export.xlsx";
        }
        // daca utilizatorul da doar nume fara cale, salvam in folderul curent
        string path = Path.Combine(Directory.GetCurrentDirectory(), fileName);
        if (CheckGroups(groups) == false)
        {
            Console.WriteLine("Export oprit: exista grupuri invalide (HostCredits < suma credite UPT).");
            return;
        }
        try
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Echivalari");
            ws.Cell(1, 1).Value = "NrCrt";
            ws.Cell(1, 2).Value = "UptName";
            ws.Cell(1, 3).Value = "UptYear";
            ws.Cell(1, 4).Value = "UptSem";
            ws.Cell(1, 5).Value = "UptCredits";
            ws.Cell(1, 6).Value = "HostName";
            ws.Cell(1, 7).Value = "HostCredits";
            ws.Cell(1, 8).Value = ""; // coloana spatiu
            ws.Cell(1, 9).Value = "Grade";

            int r = 2;
            double totalUpt = 0;
            double totalHost = 0;
            foreach (var g in groups)
            {
                bool first = true;
                totalHost += g.HostCredits;
                foreach (var c in g.UptCourses)
                {
                    ws.Cell(r, 1).Value = r - 1; // NrCrt simplu incremental
                    ws.Cell(r, 2).Value = c.Name;
                    ws.Cell(r, 3).Value = c.Year;
                    ws.Cell(r, 4).Value = c.Semester;
                    ws.Cell(r, 5).Value = c.Credits;
                    if (c.Credits.HasValue)
                    {
                        totalUpt += c.Credits.Value;
                    }
                    ws.Cell(r, 6).Value = g.HostName;
                    if (first == true)
                    {
                        ws.Cell(r, 7).Value = g.HostCredits;
                    }
                    else
                    {
                        ws.Cell(r, 7).Value = (double?)null;
                    }
                    ws.Cell(r, 9).Value = c.Grade;
                    first = false;
                    r++;
                }
            }
            // borduri pentru toate celulele cu date (fara randul total)
            var dataRange = ws.Range(1, 1, r - 1, 9);
            dataRange.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            dataRange.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            dataRange.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            dataRange.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            // centreaza numeric (NrCrt, Year, Sem, Credits, HostCredits)
            ws.Range(2, 1, r - 1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Range(2, 3, r - 1, 3).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Range(2, 4, r - 1, 4).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Range(2, 5, r - 1, 5).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Range(2, 7, r - 1, 7).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            // rand total
            r++; // rand liber intre tabel si total
            ws.Cell(r, 1).Value = "TOTAL";
            ws.Cell(r, 6).Value = "TOTAL";
            ws.Cell(r, 5).Value = totalUpt;   // sub coloana UptCredits
            ws.Cell(r, 7).Value = totalHost;  // sub HostCredits
            // elimina bordura de jos pe coloana spatiu (8) si grade (9) pentru randul total
            ws.Cell(r, 8).Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;
            var noteTotalCell = ws.Cell(r, 9);
            noteTotalCell.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;
            ws.Cell(r, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Cell(r, 5).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Cell(r, 6).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Cell(r, 7).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
            Console.WriteLine($"Export complet: {path}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Eroare la export Excel: {ex.Message}");
        }
    }

    static void ExportDocx(List<Group> groups, string[] parts)
    {
        if (CheckGroups(groups) == false)
        {
            Console.WriteLine("Export DOCX oprit: exista grupuri invalide (HostCredits < suma credite UPT).");
            return;
        }

        string templatePath = "";
        if (parts.Length > 1)
        {
            templatePath = parts[1];
        }
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        if (parts.Length > 2)
        {
            outputPath = parts[2];
        }

        if (string.IsNullOrWhiteSpace(templatePath))
        {
            templatePath = AutoFindDocx();
            if (string.IsNullOrWhiteSpace(templatePath) == false)
            {
                Console.WriteLine("Gasit automat template DOCX: " + templatePath);
            }
        }

        if (string.IsNullOrWhiteSpace(templatePath) || File.Exists(templatePath) == false)
        {
            Console.Write("Template DOCX: ");
            var input = Console.ReadLine();
            if (input != null)
            {
                templatePath = input;
            }
        }
        templatePath = templatePath.Trim().Trim('"');
        if (File.Exists(templatePath) == false)
        {
            Console.WriteLine("Template DOCX inexistent. Oprire export.");
            return;
        }

        // genereaza mai intai un Excel intermediar pe baza grupurilor editate
        var tempExcel = Path.Combine(Path.GetTempPath(), "export-echivalari-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            ExportExcel(groups, new[] { "export-excel", tempExcel });
            Console.WriteLine($"Excel intermediar generat: {tempExcel}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Nu am reusit sa generezi Excel intermediar pentru DOCX: {ex.Message}");
            return;
        }

        try
        {
            // copiem template-ul in output si il modificam
            File.Copy(templatePath, outputPath, true);
            using var wordDoc = WordprocessingDocument.Open(outputPath, true);
            var body = wordDoc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                Console.WriteLine("Nu am putut citi body din DOCX.");
                return;
            }

            var table = body.Elements<Table>().FirstOrDefault();
            if (table == null)
            {
                Console.WriteLine("Template-ul nu contine tabele.");
                return;
            }

            // pastram header-ul (primul rand) si il clonam pentru stil
            var headerRow = table.Elements<TableRow>().FirstOrDefault();
            if (headerRow == null)
            {
                Console.WriteLine("Tabel fara randuri.");
                return;
            }

            // stergem toate randurile de date
            var rows = table.Elements<TableRow>().ToList();
            foreach (var row in rows.Skip(1))
                row.Remove();

            int nr = 1;
            foreach (var g in groups)
            {
                int groupNr = nr;
                nr++;
                var groupRows = new List<TableRow>();
                bool first = true;
                foreach (var c in g.UptCourses)
                {
                    var newRow = (TableRow)headerRow.CloneNode(true);
                    var cells = newRow.Elements<TableCell>().ToList();
                    if (cells.Count < 5)
                    {
                        Console.WriteLine("Tabelul trebuie sa aiba cel putin 5 coloane.");
                        return;
                    }
                    SetCellText(cells[0], groupNr.ToString());
                    SetCellText(cells[1], g.HostName);
                    SetCellText(cells[2], c.Name);
                    double creditsValue = 0;
                    if (c.Credits.HasValue)
                    {
                        creditsValue = c.Credits.Value;
                    }
                    SetCellText(cells[3], creditsValue.ToString(CultureInfo.InvariantCulture));
                    if (c.Grade == null)
                    {
                        SetCellText(cells[4], string.Empty);
                    }
                    else
                    {
                        SetCellText(cells[4], c.Grade);
                    }
                    groupRows.Add(newRow);
                    table.AppendChild(newRow);
                    first = false;
                }
                if (groupRows.Count > 1)
                {
                    // aplica vMerge pe NrCrt (col 0) si HostName (col 1)
                    SetVerticalMerge(groupRows[0].Elements<TableCell>().ElementAt(0), MergedCellValues.Restart);
                    SetVerticalMerge(groupRows[0].Elements<TableCell>().ElementAt(1), MergedCellValues.Restart);
                    for (int i = 1; i < groupRows.Count; i++)
                    {
                        var cells = groupRows[i].Elements<TableCell>().ToList();
                        SetCellText(cells[0], "");
                        SetCellText(cells[1], "");
                        SetVerticalMerge(cells[0], MergedCellValues.Continue);
                        SetVerticalMerge(cells[1], MergedCellValues.Continue);
                    }
                }
            }

            wordDoc.MainDocumentPart.Document.Save();
            Console.WriteLine($"Export DOCX complet: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Eroare la export DOCX: {ex.Message}");
        }
    }

    static void SplitUpt(List<Group> groups, string[] parts)
    {
        // Muta o disciplina UPT dintr-un grup sursa intr-un grup tinta deja existent.
        if (parts.Length < 4 ||
            int.TryParse(parts[1], out var srcGid) == false ||
            int.TryParse(parts[2], out var idx) == false ||
            int.TryParse(parts[3], out var dstGid) == false)
        {
            Console.WriteLine("Folosire: split-upt <groupId_sursa> <uptIndex> <groupId_tinta>");
            return;
        }

        var src = groups.FirstOrDefault(g => g.Id == srcGid);
        var dst = groups.FirstOrDefault(g => g.Id == dstGid);
        if (src == null || dst == null)
        {
            Console.WriteLine("GroupId sursa sau tinta inexistent.");
            return;
        }
        if (TryGetCourse(src, idx, out var course) == false)
        {
            Console.WriteLine("uptIndex invalid pentru grupul sursa.");
            return;
        }

        // Muta disciplina
        dst.UptCourses.Add(course);
        src.UptCourses.RemoveAt(idx - 1);
        Console.WriteLine($"Mutat UPT[{idx}] din grup {srcGid} in grup {dstGid}.");
    }

    static void AddGroup(List<Group> groups, string[] parts)
    {
        if (parts.Length < 3 || double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out var credits) == false)
        {
            Console.WriteLine("Folosire: add-group <hostName> <hostCredits>");
            return;
        }
        string hostName = parts[1];
        int newId = groups.Count == 0 ? 1 : groups.Max(g => g.Id) + 1;
        groups.Add(new Group(newId, hostName, credits));
        Console.WriteLine($"Creat grup gol {newId} cu host '{hostName}' ({credits}).");
    }

    static bool TryGetCourse(Group? group, int idx, out Course course)
    {
        course = new Course();
        if (group == null || idx < 1 || idx > group.UptCourses.Count)
            return false;
        course = group.UptCourses[idx - 1];
        return true;
    }

    static int? ReadIntOrNull(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;
        if (cell.TryGetValue<int>(out var v)) return v;
        return null;
    }

    static string AutoFindExcel()
    {
        var dir = Directory.GetCurrentDirectory();
        var matches = Directory.GetFiles(dir, "tabel de echivalare discipline*.xlsx", SearchOption.TopDirectoryOnly);
        if (matches.Length > 0)
        {
            return matches[0];
        }
        return string.Empty;
    }

    static string AutoFindDocx()
    {
        var dir = Directory.GetCurrentDirectory();
        var matches = Directory.GetFiles(dir, "*.docx", SearchOption.TopDirectoryOnly);
        var pv = matches.FirstOrDefault(f => Path.GetFileName(f).Contains("PV", StringComparison.OrdinalIgnoreCase));
        if (pv != null)
        {
            return pv;
        }
        if (matches.Length > 0)
        {
            return matches[0];
        }
        return string.Empty;
    }

    static double SumCredits(List<Course> courses)
    {
        double total = 0;
        foreach (var c in courses)
        {
            if (c.Credits.HasValue)
            {
                total += c.Credits.Value;
            }
        }
        return total;
    }

    static void SetCellText(TableCell cell, string text)
    {
        cell.RemoveAllChildren<Paragraph>();
        if (text == null)
        {
            text = string.Empty;
        }
        var p = new Paragraph(new Run(new Text(text)));
        cell.AppendChild(p);
    }

    static void SetVerticalMerge(TableCell cell, MergedCellValues mergeVal)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr == null)
        {
            tcPr = new TableCellProperties();
            cell.PrependChild(tcPr);
        }
        var vMerge = tcPr.GetFirstChild<VerticalMerge>();
        if (vMerge == null)
        {
            vMerge = new VerticalMerge();
            tcPr.AppendChild(vMerge);
        }
        vMerge.Val = mergeVal;
    }

    static void PrintHelp()
    {
        Console.WriteLine("Comenzi:");
        Console.WriteLine("  list");
        Console.WriteLine("  check");
        Console.WriteLine("  set-host-name <groupId> <nume>");
        Console.WriteLine("  set-host-credits <groupId> <credite>");
        Console.WriteLine("  set-upt-name <groupId> <uptIndex> <nume>");
        Console.WriteLine("  set-upt-credits <groupId> <uptIndex> <credite>");
        Console.WriteLine("  set-upt-grade <groupId> <uptIndex> <grade>");
        Console.WriteLine("  add-upt <groupId> <nume> <credite> [an] [sem] [grade]");
        Console.WriteLine("  remove-upt <groupId> <uptIndex>");
        Console.WriteLine("  split-upt <groupId_sursa> <uptIndex> <groupId_tinta>  (muta disciplina UPT intr-un grup tinta existent)");
        Console.WriteLine("  add-group <hostName> <hostCredits>");
        Console.WriteLine("  export-excel [numeFisier]  (salveaza in folderul curent)");
        Console.WriteLine("  export-docx <templateDocx> [outputDocx]");
        Console.WriteLine("  help");
        Console.WriteLine("  exit");
    }
}
