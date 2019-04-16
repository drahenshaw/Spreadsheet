// <copyright file="SpreadsheetCell.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System.ComponentModel;
using System.Xml;

namespace CptS321
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    public delegate void DependencyChangedEventHandler(object sender);

    /// <summary>
    /// Abstract Class for a Spreadsheet Cell with NotifyPropertyChanged Interface
    /// </summary>
    public abstract class SpreadsheetCell : INotifyPropertyChanged
    {
        /// <summary>
        /// Number of Rows
        /// </summary>
        protected uint   rowIndex;

        /// <summary>
        /// Number of Columns
        /// </summary>
        protected uint   columnIndex;

        /// <summary>
        /// Text of Cell
        /// </summary>
        protected string cellText;

        /// <summary>
        /// Value of Cell
        /// </summary>
        protected string cellValue;

        /// <summary>
        /// Set the Delegate for Handling Events
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        /// <summary>
        /// Occurs when cell delegate changes
        /// </summary>
        public event DependencyChangedEventHandler DependencyChanged;

        /// <summary>
        /// Background Color for the Cell
        /// </summary>
        protected uint cellColor;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        public void OnDependencyChanged(object sender)
        {
            this.OnPropertyChanged("CellText");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        protected void OnPropertyChanged(string name)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            this.DependencyChanged?.Invoke(this);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetCell"/> class.
        /// Initalizes SpreadSheetCell with Row Column Indices
        /// </summary>
        /// <param name="row">Row</param>
        /// <param name="column">Column</param>
        public SpreadsheetCell(uint row, uint column)
        {
            this.rowIndex = row;
            this.columnIndex = column;
            this.cellText = string.Empty;
            this.cellValue = string.Empty;
            this.BGColor = 0xFFFFFFFF;
        }

        /// <summary>
        /// Gets ReadOnly Property for RowIndex
        /// </summary>
        public uint RowIndex
        {
            get { return this.rowIndex; }
        }

        /// <summary>
        /// Gets ReadOnly Property for Column Index
        /// </summary>
        public uint ColumnIndex
        {
            get { return this.columnIndex; }
        }

        /// <summary>
        /// Gets or sets Property for CellText - Raises Event on Set
        /// </summary>
        public string CellText
        {
            get { return this.cellText; }

            set
            {
                // If Text is Unchaged Return
                if (this.cellText == value) { return; }

                // Otherwise Update Text
                this.cellText = value;

                // Raise Property Changed Event
                this.OnPropertyChanged("CellText");
            }
        }

        /// <summary>
        /// Gets or sets Property for CellValue: Publically Accessible, Protected Internal Mutatable
        /// </summary>
        public string CellValue
        {
            // Allow All to Obtain the Value
            get { return this.cellValue; }

            // Only Classes from within the DLL May Mutate the Value
            protected internal set
            {
                // If Value is Unchanged Return
                if (this.cellValue == value) { return; }

                // Otherwise Update Value
                this.cellValue = value;

                // Raise Property Changed Event
                this.OnPropertyChanged("CellValue");
            }
        }

        /// <summary>
        /// Gets or sets Property for Background Color
        /// </summary>
        public uint BGColor
        {
            get { return this.cellColor; }

            set
            {
                if (cellColor == value)
                {
                    return;
                }

                this.cellColor = value;
                this.OnPropertyChanged("BGColor");
            }
        }

        /// <summary>
        /// Saves the state of the cell using XML
        /// </summary>
        /// <param name="writer">XML writer</param>
        public void SaveCell(XmlWriter writer)
        {
            // Only save if the cell does not contain default values
            if (this.CellText != string.Empty || this.BGColor != 0xFFFFFFFF)
            {
                // Convert the Cell name
                string name = (char)(this.ColumnIndex + (int)'A') + (this.RowIndex + 1).ToString();

                // Write desired cell attributes to XML
                writer.WriteStartElement("SpreadsheetCell");
                writer.WriteAttributeString("Name", name);
                writer.WriteAttributeString("Text", this.CellText);
                writer.WriteAttributeString("BGColor", this.BGColor.ToString());
                writer.WriteEndElement();
            }           
        }
    }
}
