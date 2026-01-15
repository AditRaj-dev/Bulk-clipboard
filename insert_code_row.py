from excel_manager import get_workbook, get_sheet
from config import CODE_COLUMN, FONT_NAME, EXCEL_FILE
from openpyxl.styles import Alignment, Font

sheet_name = "DSA Pre Mid Term"
row = 24
code = '''#include <stdio.h>

int main() {
    int n;
    scanf("%d", &n); // Read the size of the array

    int arr[n];

    // Read array elements
    for (int i = 0; i < n; i++) {
        scanf("%d", &arr[i]);
    }

    // Selection Sort
    for (int i = 0; i < n - 1; i++) {
        int min_idx = i;
        for (int j = i + 1; j < n; j++) {
            if (arr[j] < arr[min_idx]) {
                min_idx = j;
            }
        }
        // Swap arr[i] and arr[min_idx]
        int temp = arr[i];
        arr[i] = arr[min_idx];
        arr[min_idx] = temp;
    }

    // Print sorted array
    for (int i = 0; i < n; i++) {
        printf("%d", arr[i]);
        if (i != n - 1) printf(" ");
    }
    printf("\n");

    return 0;
}
'''

wb, ws = get_sheet(sheet_name)
cell = ws.cell(row=row, column=CODE_COLUMN)
cell.value = code
cell.alignment = Alignment(wrap_text=True)
cell.font = Font(name=FONT_NAME)
wb.save(EXCEL_FILE)
print(f"Inserted code into {EXCEL_FILE} sheet '{sheet_name}' row {row}")
