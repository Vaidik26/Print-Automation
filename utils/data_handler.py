"""
Data Handler Module
Handles reading and processing data from CSV and Excel files.
"""

import pandas as pd
from io import BytesIO
from typing import List, Dict, Any, Tuple


class DataHandler:
    """Handle data file reading and processing."""

    SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}

    def __init__(self, file_bytes: bytes, filename: str):
        """
        Initialize with file bytes and filename.

        Args:
            file_bytes: The data file as bytes
            filename: Original filename (used to determine file type)
        """
        self.file_bytes = file_bytes
        self.filename = filename.lower()
        self.df = None
        self.columns = []
        self._load_data()

    def _load_data(self):
        """Load data from the file into a pandas DataFrame."""
        try:
            if self.filename.endswith(".csv"):
                # Try different encodings
                for encoding in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        self.df = pd.read_csv(
                            BytesIO(self.file_bytes), encoding=encoding
                        )
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError(
                        "Could not decode CSV file with any supported encoding"
                    )

            elif self.filename.endswith(".xlsx"):
                self.df = pd.read_excel(BytesIO(self.file_bytes), engine="openpyxl")

            elif self.filename.endswith(".xls"):
                self.df = pd.read_excel(BytesIO(self.file_bytes), engine="xlrd")

            else:
                raise ValueError(
                    f"Unsupported file format. Supported formats: {self.SUPPORTED_EXTENSIONS}"
                )

            # Store column names
            self.columns = list(self.df.columns)

            # Clean column names (strip whitespace)
            self.df.columns = [str(col).strip() for col in self.df.columns]
            self.columns = list(self.df.columns)

        except Exception as e:
            raise ValueError(f"Error loading data file: {str(e)}")

    def get_columns(self) -> List[str]:
        """
        Get list of column names from the data file.

        Returns:
            List of column names
        """
        return self.columns

    def get_row_count(self) -> int:
        """
        Get the number of data rows.

        Returns:
            Number of rows in the data
        """
        return len(self.df) if self.df is not None else 0

    def get_preview(self, num_rows: int = 5) -> pd.DataFrame:
        """
        Get a preview of the data.

        Args:
            num_rows: Number of rows to preview

        Returns:
            DataFrame with preview rows
        """
        if self.df is None:
            return pd.DataFrame()
        return self.df.head(num_rows)

    def get_data_as_dicts(
        self, column_mapping: Dict[str, str] = None
    ) -> List[Dict[str, Any]]:
        """
        Get all data rows as a list of dictionaries.

        Args:
            column_mapping: Optional mapping of placeholder names to column names
                           {placeholder: column_name}

        Returns:
            List of dictionaries, one per row
        """
        if self.df is None:
            return []

        # Convert DataFrame to list of dicts
        rows = self.df.to_dict("records")

        # If no mapping provided, return as-is
        if not column_mapping:
            # Convert all values to strings
            return [{k: self._format_value(v) for k, v in row.items()} for row in rows]

        # Apply column mapping
        mapped_rows = []
        for row in rows:
            mapped_row = {}
            for placeholder, column_name in column_mapping.items():
                if column_name in row:
                    mapped_row[placeholder] = self._format_value(row[column_name])
                else:
                    mapped_row[placeholder] = ""
            # Also include original columns for filename generation
            for col in row:
                if col not in mapped_row:
                    mapped_row[col] = self._format_value(row[col])
            mapped_rows.append(mapped_row)

        return mapped_rows

    def _format_value(self, value: Any) -> str:
        """
        Format a value for insertion into a document.

        Args:
            value: The value to format

        Returns:
            Formatted string value
        """
        if pd.isna(value):
            return ""

        # Handle different types
        if isinstance(value, float):
            # Check if it's a whole number
            if value.is_integer():
                return str(int(value))
            return str(value)

        return str(value)

    def validate_mapping(
        self, column_mapping: Dict[str, str]
    ) -> Tuple[bool, List[str]]:
        """
        Validate that all mapped columns exist in the data.

        Args:
            column_mapping: Mapping of placeholder to column names

        Returns:
            Tuple of (is_valid, list_of_missing_columns)
        """
        missing = []
        for placeholder, column in column_mapping.items():
            if column and column not in self.columns:
                missing.append(column)

        return len(missing) == 0, missing

    def get_unique_values(self, column: str, limit: int = 10) -> List[str]:
        """
        Get unique values from a column (for preview purposes).

        Args:
            column: Column name
            limit: Maximum number of unique values to return

        Returns:
            List of unique values
        """
        if self.df is None or column not in self.columns:
            return []

        unique = self.df[column].dropna().unique()
        return [self._format_value(v) for v in unique[:limit]]
