using ClosedXML.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal sealed class XLRow : XLRangeBase, IXLRow
    {
        #region Private fields

        /// <summary>
        /// Don't use directly, use properties.
        /// </summary>
        private XlRowFlags _flags;
        private Double _height;
        private Int32 _outlineLevel;

        #endregion Private fields

        #region Constructor

        /// <summary>
        /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRow(XLWorksheet worksheet, Int32 row)
            : base(XLRangeAddress.EntireRow(worksheet, row), worksheet.StyleValue)
        {
            SetRowNumber(row);

            _height = worksheet.RowHeight;
        }

        #endregion Constructor

        public override XLRangeType RangeType
        {
            get { return XLRangeType.Row; }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                int row = RowNumber();

                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell;
            }
        }

        public Boolean Collapsed
        {
            get => _flags.HasFlag(XlRowFlags.Collapsed);
            set
            {
                if (value)
                    _flags |= XlRowFlags.Collapsed;
                else
                    _flags &= ~XlRowFlags.Collapsed;
            }
        }

        /// <summary>
        /// Distance in pixels from the bottom of the cells in the current row to the typographical
        /// baseline of the cell content if, hypothetically, the zoom level for the sheet containing
        /// this row is 100 percent and the cell has bottom-alignment formatting.
        /// </summary>
        /// <remarks>
        /// If the attribute is set, it sets customHeight to true even if the customHeight is explicitly
        /// set to false. Custom height means no auto-sizing by Excel on load, so if row has this
        /// attribute, it stops Excel from auto-sizing the height of a row to fit the content on load.
        /// </remarks>
        public Double? DyDescent { get; set; }

        /// <summary>
        /// Should cells in the row display phonetic? This doesn't actually affect whether the phonetic are
        /// shown in the row, that depends entirely on the <see cref="IXLCell.ShowPhonetic"/> property
        /// of a cell. This property determines whether a new cell in the row will have it's phonetic turned on
        /// (and also the state of the "Show or hide phonetic" in Excel when whole row is selected).
        /// Default is <c>false</c>.
        /// </summary>
        public Boolean ShowPhonetic
        {
            get => _flags.HasFlag(XlRowFlags.ShowPhonetic);
            set
            {
                if (value)
                    _flags |= XlRowFlags.ShowPhonetic;
                else
                    _flags &= ~XlRowFlags.ShowPhonetic;
            }
        }

        public Boolean Loading
        {
            get => _flags.HasFlag(XlRowFlags.Loading);
            set
            {
                if (value)
                    _flags |= XlRowFlags.Loading;
                else
                    _flags &= ~XlRowFlags.Loading;
            }
        }

        /// <summary>
        /// Does row have an individual height or is it derived from the worksheet <see cref="XLWorksheet.RowHeight"/>?
        /// </summary>
        public Boolean HeightChanged
        {
            get => _flags.HasFlag(XlRowFlags.HeightChanged);
            private set
            {
                if (value)
                    _flags |= XlRowFlags.HeightChanged;
                else
                    _flags &= ~XlRowFlags.HeightChanged;
            }
        }

        #region IXLRow Members

        public Double Height
        {
            get { return _height; }
            set
            {
                if (!Loading)
                    HeightChanged = true;

                _height = value;
            }
        }

        IXLCells IXLRow.Cells(String cellsInRow) => Cells(cellsInRow);

        IXLCells IXLRow.Cells(Int32 firstColumn, Int32 lastColumn) => Cells(firstColumn, lastColumn);

        public void ClearHeight()
        {
            Height = Worksheet.RowHeight;
            HeightChanged = false;
        }

        public void Delete()
        {
            int rowNumber = RowNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            Worksheet.DeleteRow(rowNumber);
        }

        public new IXLRows InsertRowsBelow(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            var asRange = Worksheet.Row(rowNum).AsRange();
            asRange.InsertRowsBelowVoid(true, numberOfRows);

            var newRows = Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);

            CopyRows(newRows);

            return newRows;
        }

        private void CopyRows(IXLRows newRows)
        {
            foreach (var newRow in newRows)
            {
                var internalRow = Worksheet.Internals.RowsCollection[newRow.RowNumber()];
                internalRow._height = Height;
                internalRow.InnerStyle = InnerStyle;
                internalRow.Collapsed = Collapsed;
                internalRow.IsHidden = IsHidden;
                internalRow._outlineLevel = OutlineLevel;
            }
        }

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
            if (rowNum > 1)
            {
                return Worksheet.Row(rowNum - 1).InsertRowsBelow(numberOfRows);
            }

            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            var asRange = Worksheet.Row(rowNum).AsRange();
            asRange.InsertRowsAboveVoid(true, numberOfRows);

            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLCell Cell(Int32 columnNumber)
        {
            return Cell(1, columnNumber);
        }

        public override XLCell Cell(String columnLetter)
        {
            return Cell(1, columnLetter);
        }

        IXLCell IXLRow.Cell(string columnLetter)
        {
            return Cell(columnLetter);
        }

        public override IXLCells Cells()
        {
            return Cells(true, XLCellsUsedOptions.All);
        }

        public override XLCells Cells(Boolean usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, XLCellsUsedOptions.AllContents);
            else
                return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
        }

        public override XLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInRow.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public XLCells Cells(Int32 firstColumn, Int32 lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(String firstColumn, String lastColumn)
        {
            return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRow AdjustToContents(Int32 startColumn)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            return AdjustToContents(startColumn, endColumn, 0, Double.MaxValue);
        }

        public IXLRow AdjustToContents(Double minHeight, Double maxHeight)
        {
            return AdjustToContents(1, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeightPt, Double maxHeightPt)
        {
            var engine = Worksheet.Workbook.GraphicEngine;
            var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);

            var rowHeightPx = CalculateMinRowHeight(startColumn, endColumn, engine, dpi);

            var rowHeightPt = XLHelper.PixelsToPoints(rowHeightPx, dpi.Y);
            if (rowHeightPt <= 0)
                rowHeightPt = Worksheet.RowHeight;

            if (minHeightPt > rowHeightPt)
                rowHeightPt = minHeightPt;

            if (maxHeightPt < rowHeightPt)
                rowHeightPt = maxHeightPt;

            Height = rowHeightPt;

            return this;
        }

        private int CalculateMinRowHeight(int startColumn, int endColumn, IXLGraphicEngine engine, Dpi dpi)
        {
            var glyphs = new List<GlyphBox>();
            var rowHeightPx = 0;
            foreach (var cell in Row(startColumn, endColumn).CellsUsed().Cast<XLCell>())
            {
                // Clear maintains capacity -> reduce need for GC
                glyphs.Clear();

                if (cell.IsMerged())
                    continue;

                var cellHeightPx = (int)Math.Ceiling(GetContentHeight(cell));
                rowHeightPx = Math.Max(cellHeightPx, rowHeightPx);
            }

            return rowHeightPx;
        }

        /// <summary>
        /// Calculates the content height for a specific cell taking into account text wrapping and column width.
        /// </summary>
        /// <param name="cell">The cell to calculate content height for.</param>
        /// <returns>The content height in pixels.</returns>
        public static double GetContentHeight(XLCell cell)
        {
            ExtractCellData(cell, out var glyphs, out var textRotation, out var colWidthPx, out var colHeightPx, out var spaceIndices, out var lineSpacing);

            return textRotation switch
            {
                0 => CalculateHorizontalTextHeight(glyphs, colWidthPx, spaceIndices, lineSpacing),
                255 => CalculateVerticalTextHeight(glyphs, colHeightPx, spaceIndices, lineSpacing),
                _ => CalculateRotatedTextHeight(glyphs, textRotation)
            };
        }

        /// <summary>
        /// Extracts necessary data from the cell for height calculation.
        /// </summary>
        private static void ExtractCellData(XLCell cell, out List<GlyphBox> glyphs, out int textRotation, out int colWidthPx, out int colHeightPx, out HashSet<int> spaceIndices, out double lineSpacing)
        {
            var engine = cell.Worksheet.Workbook.GraphicEngine;
            var dpi = new Dpi(cell.Worksheet.Workbook.DpiX, cell.Worksheet.Workbook.DpiY);
            glyphs = new List<GlyphBox>();

            // Get glyph boxes for the cell
            cell.GetGlyphBoxes(engine, dpi, glyphs);

            // Get text rotation from cell style
            textRotation = cell.Style.Alignment.TextRotation;

            // Get column width in pixels
            colWidthPx = GetColumnWidthInPixels(cell, engine, dpi);
            colHeightPx = GetColumnHeightInPixels(cell, engine, dpi);

            // Calculate proper line spacing from font metrics
            lineSpacing = CalculateLineSpacing(cell, engine, dpi);

            // Pre-calculate space indices for efficient word boundary detection
            spaceIndices = new HashSet<int>();
            for (int i = 0; i < glyphs.Count; i++)
            {
                var glyph = glyphs[i];
                // Detect spaces based on glyph properties
                // Spaces typically have advance width roughly 1/4 to 1/3 of the line height
                bool isSpace = glyph.AdvanceWidth > 0 &&
                              glyph.AdvanceWidth < glyph.LineHeight * 0.4f &&
                              glyph.AdvanceWidth >= glyph.LineHeight * 0.15f;
                if (isSpace)
                {
                    spaceIndices.Add(i);
                }
            }
        }

        /// <summary>
        /// Calculates the column width in pixels for the given cell.
        /// </summary>
        private static int GetColumnWidthInPixels(XLCell cell, IXLGraphicEngine engine, Dpi dpi)
        {
            // Get the column width in Excel's "Number of Characters" unit
            var column = cell.Worksheet.Column(cell.Address.ColumnNumber);
            var columnWidthNoC = column.Width;

            // Calculate Maximum Digit Width (MDW) for conversion
            var mdw = (int)Math.Round(engine.GetMaxDigitWidth(cell.Worksheet.Workbook.Style.Font, dpi.X));

            // Convert from NoC to pixels using Excel's formula
            var columnWidthPx = (int)Math.Ceiling(XLHelper.NoCToPixels(columnWidthNoC, mdw));

            // Use a reasonable minimum width if column width is too small
            return Math.Max(columnWidthPx, 20); // Minimum 20 pixels
        }

        /// <summary>
        /// Calculates the column height in pixels for the given cell.
        /// </summary>
        private static int GetColumnHeightInPixels(XLCell cell, IXLGraphicEngine engine, Dpi dpi)
        {
            // Get the row height in points
            var row = cell.WorksheetRow();
            var rowHeightPt = row.Height;

            // Convert row height from points to pixels
            var rowHeightPx = (int)Math.Ceiling(XLHelper.PointsToPixels(rowHeightPt, dpi.Y));

            // Use a reasonable minimum height if row height is too small
            return Math.Max(rowHeightPx, 20); // Minimum 20 pixels (about 15 points)
        }

        /// <summary>
        /// Calculates proper line spacing based on font metrics.
        /// </summary>
        private static double CalculateLineSpacing(XLCell cell, IXLGraphicEngine engine, Dpi dpi)
        {
            // Get the font's proper text height which includes line spacing
            var fontTextHeight = engine.GetTextHeight(cell.Style.Font, dpi.Y);

            // Get the actual glyph line height (EmSize + Descent)
            var glyphLineHeight = cell.Style.Font.FontSize / 72d * dpi.Y + engine.GetDescent(cell.Style.Font, dpi.Y);

            // Line spacing is the difference between font's text height and glyph line height
            // This represents the natural spacing that should exist between lines
            return Math.Max(0, fontTextHeight - glyphLineHeight);
        }

        /// <summary>
        /// Calculates height for horizontal text with word wrapping support.
        /// </summary>
        private static double CalculateHorizontalTextHeight(List<GlyphBox> glyphs, int colWidthPx, HashSet<int> spaceIndices, double lineSpacing)
        {
            if (colWidthPx <= 0 || glyphs.Count == 0)
            {
                return CalculateFallbackHeight(glyphs, lineSpacing);
            }

            var textHeight = 0d;
            var lineMaxHeight = 0d;
            var currentLineWidth = 0d;
            var currentWordWidth = 0d;
            var currentWordGlyphCount = 0;
            var isFirstWordOfLine = true;
            var lineCount = 0;

            for (int i = 0; i < glyphs.Count; i++)
            {
                var glyph = glyphs[i];

                if (glyph.IsLineBreak)
                {
                    textHeight += lineMaxHeight;
                    lineMaxHeight = 0d;
                    currentLineWidth = 0d;
                    currentWordWidth = 0d;
                    currentWordGlyphCount = 0;
                    isFirstWordOfLine = true;

                    lineCount++;
                    continue;
                }

                var glyphHeight = glyph.LineHeight;
                lineMaxHeight = Math.Max(glyphHeight, lineMaxHeight);

                bool isWordEnd = spaceIndices.Contains(i) || i == glyphs.Count - 1;

                currentWordWidth += glyph.AdvanceWidth;
                currentWordGlyphCount++;

                if (isWordEnd)
                {
                    bool wasWrapped = false;

                    // End of word - check if we need to wrap
                    if (currentLineWidth + currentWordWidth > colWidthPx)
                    {
                        if (isFirstWordOfLine)
                        {

                            // First word doesn't fit - wrap on character boundary
                            var remainingWidth = colWidthPx - currentLineWidth;
                            if (remainingWidth > 0 && currentWordGlyphCount > 1)
                            {
                                // Estimate how many characters fit
                                var avgGlyphWidth = currentWordWidth / currentWordGlyphCount;
                                var fittingGlyphs = Math.Max(1, (int)(remainingWidth / avgGlyphWidth));

                                if (fittingGlyphs < currentWordGlyphCount)
                                {
                                    // Split the word - start new line with remaining part
                                    textHeight += lineMaxHeight;
                                    lineMaxHeight = glyphHeight;
                                    currentLineWidth = currentWordWidth - (fittingGlyphs * avgGlyphWidth);
                                }
                                else
                                {
                                    // Whole word fits
                                    currentLineWidth += currentWordWidth;
                                }
                            }
                            else
                            {
                                // Can't fit anything, start new line
                                textHeight += lineMaxHeight;
                                lineMaxHeight = glyphHeight;
                                currentLineWidth = currentWordWidth;
                            }
                        }
                        else
                        {
                            // Not first word - wrap on word boundary
                            textHeight += lineMaxHeight;
                            lineMaxHeight = glyphHeight;
                            currentLineWidth = currentWordWidth;
                        }
                        isFirstWordOfLine = false;
                        wasWrapped = true;
                    }
                    else
                    {
                        // Word fits on current line
                        currentLineWidth += currentWordWidth;
                        isFirstWordOfLine = false;
                    }

                    currentWordWidth = 0d;
                    currentWordGlyphCount = 0;

                    if (wasWrapped)
                        lineCount++;
                }
            }

            // Add height of the final line if it has content
            if (lineMaxHeight > 0)
            {
                textHeight += lineMaxHeight;
                lineCount++;
            }

            // Add proper line spacing between lines based on font metrics
            if (lineCount > 1)
            {
                textHeight += lineSpacing * (lineCount - 1); // Add spacing between lines
            }

            return textHeight;
        }

        /// <summary>
        /// Calculates the height needed for vertical text (rotation 255).
        /// For vertical text, we wrap based on available height and accumulate glyph widths.
        /// </summary>
        private static double CalculateVerticalTextHeight(List<GlyphBox> glyphs, int colHeightPx, HashSet<int> spaceIndices, double lineSpacing)
        {
            if (!glyphs.Any())
                return 0;

            var availableHeight = colHeightPx;
            var currentHeight = 0.0;
            var totalWidth = 0.0;
            var columnCount = 1;
            var maxGlyphWidth = 0.0;
            var currentWordHeight = 0.0;
            var currentWordMaxWidth = 0.0;
            var currentWordGlyphCount = 0;
            var isFirstWordOfColumn = true;

            for (int i = 0; i < glyphs.Count; i++)
            {
                var glyph = glyphs[i];

                // Check for explicit line breaks (\n) - for vertical text, this starts a new column
                if (ProcessVerticalLineBreak(glyph, ref currentHeight, ref totalWidth, ref columnCount, ref maxGlyphWidth))
                {
                    // Reset word tracking after line break
                    currentWordHeight = 0.0;
                    currentWordMaxWidth = 0.0;
                    currentWordGlyphCount = 0;
                    isFirstWordOfColumn = true;
                    continue;
                }

                var glyphHeight = glyph.LineHeight;
                var glyphWidth = glyph.AdvanceWidth;

                bool isWordEnd = spaceIndices.Contains(i) || i == glyphs.Count - 1;

                // Accumulate word dimensions
                currentWordHeight += glyphHeight;
                currentWordMaxWidth = Math.Max(currentWordMaxWidth, glyphWidth);
                currentWordGlyphCount++;

                if (isWordEnd)
                {
                    // End of word - check if we need to wrap to next column
                    if (currentHeight + currentWordHeight > availableHeight && currentHeight > 0)
                    {
                        if (isFirstWordOfColumn)
                        {
                            // First word doesn't fit - split it (character-level wrapping for vertical text)
                            var remainingHeight = availableHeight - currentHeight;
                            if (remainingHeight > 0 && currentWordGlyphCount > 1)
                            {
                                // Estimate how many characters fit
                                var avgGlyphHeight = currentWordHeight / currentWordGlyphCount;
                                var fittingGlyphs = Math.Max(1, (int)(remainingHeight / avgGlyphHeight));

                                if (fittingGlyphs < currentWordGlyphCount)
                                {
                                    // Split the word - start new column with remaining part
                                    totalWidth += maxGlyphWidth;
                                    columnCount++;
                                    currentHeight = currentWordHeight - (fittingGlyphs * avgGlyphHeight);
                                    maxGlyphWidth = currentWordMaxWidth;
                                }
                                else
                                {
                                    // Whole word fits
                                    currentHeight += currentWordHeight;
                                    maxGlyphWidth = Math.Max(maxGlyphWidth, currentWordMaxWidth);
                                }
                            }
                            else
                            {
                                // Can't fit anything, start new column
                                totalWidth += maxGlyphWidth;
                                columnCount++;
                                currentHeight = currentWordHeight;
                                maxGlyphWidth = currentWordMaxWidth;
                            }
                        }
                        else
                        {
                            // Not first word - wrap on word boundary
                            totalWidth += maxGlyphWidth;
                            columnCount++;
                            currentHeight = currentWordHeight;
                            maxGlyphWidth = currentWordMaxWidth;
                        }
                        isFirstWordOfColumn = false;
                    }
                    else
                    {
                        // Word fits in current column
                        currentHeight += currentWordHeight;
                        maxGlyphWidth = Math.Max(maxGlyphWidth, currentWordMaxWidth);
                        isFirstWordOfColumn = false;
                    }

                    // Reset word tracking
                    currentWordHeight = 0.0;
                    currentWordMaxWidth = 0.0;
                    currentWordGlyphCount = 0;
                }
            }

            // Add the width of the last column
            totalWidth += maxGlyphWidth;

            // Add proper spacing between columns based on font metrics
            // For vertical text, column spacing should be proportional to line spacing
            if (columnCount > 1)
            {
                totalWidth += lineSpacing * (columnCount - 1);
            }

            return totalWidth;
        }

        /// <summary>
        /// Processes line breaks for vertical text - creates a new column.
        /// </summary>
        private static bool ProcessVerticalLineBreak(GlyphBox glyph, ref double currentHeight, ref double totalWidth, ref int columnCount, ref double maxGlyphWidth)
        {
            // Check if this is a line break character (typically has very small AdvanceWidth and specific characteristics)
            if (glyph.AdvanceWidth < glyph.LineHeight * 0.1 && glyph.EmSize > 0)
            {
                // Start new column
                totalWidth += maxGlyphWidth;
                columnCount++;
                currentHeight = 0;
                maxGlyphWidth = 0;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Calculates height for rotated text.
        /// </summary>
        private static double CalculateRotatedTextHeight(List<GlyphBox> glyphs, int textRotation)
        {
            var width = 0d;
            var height = 0d;
            foreach (var glyph in glyphs)
            {
                width += glyph.AdvanceWidth;
                height = Math.Max(glyph.LineHeight, height);
            }

            var projectedWidth = Math.Sin(XLHelper.DegToRad(textRotation)) * width;
            var projectedHeight = Math.Cos(XLHelper.DegToRad(textRotation)) * height;
            return projectedWidth + projectedHeight;
        }

        /// <summary>
        /// Calculates fallback height when column width is not available.
        /// </summary>
        private static double CalculateFallbackHeight(List<GlyphBox> glyphs, double lineSpacing)
        {
            var fallbackHeight = 0d;
            var fallbackLineMaxHeight = 0d;
            var lineCount = 0;

            foreach (var glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    var cellHeightPx = glyph.LineHeight;
                    fallbackLineMaxHeight = Math.Max(cellHeightPx, fallbackLineMaxHeight);
                }
                else
                {
                    fallbackHeight += fallbackLineMaxHeight;
                    fallbackLineMaxHeight = 0d;
                    lineCount++;
                }
            }

            if (fallbackLineMaxHeight > 0)
            {
                fallbackHeight += fallbackLineMaxHeight;
                lineCount++;
            }

            // Add proper line spacing for fallback calculation based on font metrics
            if (lineCount > 1)
            {
                fallbackHeight += lineSpacing * (lineCount - 1);
            }

            return fallbackHeight;
        }



        public IXLRow Hide()
        {
            IsHidden = true;
            return this;
        }

        public IXLRow Unhide()
        {
            IsHidden = false;
            return this;
        }

        public Boolean IsHidden
        {
            get => _flags.HasFlag(XlRowFlags.IsHidden);
            set
            {
                if (value)
                    _flags |= XlRowFlags.IsHidden;
                else
                    _flags &= ~XlRowFlags.IsHidden;
            }
        }

        public Int32 OutlineLevel
        {
            get { return _outlineLevel; }
            set
            {
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");

                Worksheet.IncrementColumnOutline(value);
                Worksheet.DecrementColumnOutline(_outlineLevel);
                _outlineLevel = value;
            }
        }

        public IXLRow Group()
        {
            return Group(false);
        }

        public IXLRow Group(Int32 outlineLevel)
        {
            return Group(outlineLevel, false);
        }

        public IXLRow Ungroup()
        {
            return Ungroup(false);
        }

        public IXLRow Group(Boolean collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
            return this;
        }

        public IXLRow Group(Int32 outlineLevel, Boolean collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
            return this;
        }

        public IXLRow Ungroup(Boolean ungroupFromAll)
        {
            if (ungroupFromAll)
                OutlineLevel = 0;
            else
            {
                if (OutlineLevel > 0)
                    OutlineLevel -= 1;
            }
            return this;
        }

        public IXLRow Collapse()
        {
            Collapsed = true;
            return Hide();
        }

        public IXLRow Expand()
        {
            Collapsed = false;
            return Unhide();
        }

        public Int32 CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public new IXLRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false,
                                          Boolean ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        IXLRangeRow IXLRow.CopyTo(IXLCell target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Row(1);
        }

        IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Row(1);
        }

        public IXLRow CopyTo(IXLRow row)
        {
            row.Clear();
            var newRow = (XLRow)row;
            newRow._height = _height;
            newRow.HeightChanged = HeightChanged;
            newRow.InnerStyle = GetStyle();
            newRow.IsHidden = IsHidden;

            AsRange().CopyTo(row);

            return newRow;
        }

        public IXLRangeRow Row(Int32 start, Int32 end)
        {
            return Range(1, start, 1, end).Row(1);
        }

        public IXLRangeRow Row(IXLCell start, IXLCell end)
        {
            return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
        }

        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (string pair in rowPairs)
                AsRange().Rows(pair.Trim()).ForEach(retVal.Add);

            return retVal;
        }

        public IXLRow AddHorizontalPageBreak()
        {
            Worksheet.PageSetup.AddHorizontalPageBreak(RowNumber());
            return this;
        }

        public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
        {
            return Row((this as IXLRangeBase).FirstCellUsed(options),
                (this as IXLRangeBase).LastCellUsed(options));
        }

        #endregion IXLRow Members

        public override XLRange AsRange()
        {
            return Range(1, 1, 1, XLHelper.MaxColumnNumber);
        }

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            //do nothing
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            // rows are shifted by XLRowCollection
        }

        internal void SetRowNumber(Int32 row)
        {
            RangeAddress = new XLRangeAddress(
                new XLAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(Worksheet,
                              row,
                              XLHelper.MaxColumnNumber,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }

        public override XLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                var arrRange = rangeAddressStr.Split(':');
                string firstPart = arrRange[0];
                string secondPart = arrRange[1];
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
                rangeAddressToUse = FixRowAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public IXLRow AdjustToContents()
        {
            return AdjustToContents(1);
        }

        internal void SetStyleNoColumns(IXLStyle value)
        {
            InnerStyle = value;

            int row = RowNumber();
            foreach (XLCell c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                c.InnerStyle = value;
        }

        private XLRow RowShift(Int32 rowsToShift)
        {
            return Worksheet.Row(RowNumber() + rowsToShift);
        }

        #region XLRow Above

        IXLRow IXLRow.RowAbove()
        {
            return RowAbove();
        }

        IXLRow IXLRow.RowAbove(Int32 step)
        {
            return RowAbove(step);
        }

        public XLRow RowAbove()
        {
            return RowAbove(1);
        }

        public XLRow RowAbove(Int32 step)
        {
            return RowShift(step * -1);
        }

        #endregion XLRow Above

        #region XLRow Below

        IXLRow IXLRow.RowBelow()
        {
            return RowBelow();
        }

        IXLRow IXLRow.RowBelow(Int32 step)
        {
            return RowBelow(step);
        }

        public XLRow RowBelow()
        {
            return RowBelow(1);
        }

        public XLRow RowBelow(Int32 step)
        {
            return RowShift(step);
        }

        #endregion XLRow Below

        public override Boolean IsEmpty()
        {
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public override Boolean IsEmpty(XLCellsUsedOptions options)
        {
            if (options.HasFlag(XLCellsUsedOptions.NormalFormats) &&
                !StyleValue.Equals(Worksheet.StyleValue))
                return false;

            return base.IsEmpty(options);
        }

        public override Boolean IsEntireRow()
        {
            return true;
        }

        public override Boolean IsEntireColumn()
        {
            return false;
        }


        /// <summary>
        /// Flag enum to save space, instead of wasting byte for each flag.
        /// </summary>
        [Flags]
        private enum XlRowFlags : byte
        {
            Collapsed = 1,
            IsHidden = 2,
            ShowPhonetic = 4,
            HeightChanged = 8,
            Loading = 16
        }
    }
}
