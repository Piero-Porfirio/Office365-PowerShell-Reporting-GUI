
'use client';

import React, { useState, useMemo, useCallback } from 'react';
import {
  Table,
  TableHeader,
  TableBody,
  TableFooter,
  TableHead,
  TableRow,
  TableCell,
  TableCaption,
} from '@/components/ui/table';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { ArrowUpDown, Download, Mail, Send, Inbox } from 'lucide-react';
import type { SharedMailboxTrafficStats } from '@/types/reporting'; // Keep specific type for potential overrides
import { saveAs } from 'file-saver'; // Using file-saver for CSV export

// Generic type for table data
type TableDataRow = Record<string, any>;

// Define sort key based on the generic row type
type SortKey<T extends TableDataRow> = keyof T | null;
type SortDirection = 'asc' | 'desc';

interface ReportTableProps {
  data: TableDataRow[]; // Accept generic data
  caption?: string;
}

// Function to generate headers dynamically from data keys
const generateHeaders = (data: TableDataRow[]): { key: string; label: string; icon?: React.ElementType }[] => {
  if (!data || data.length === 0) return [];

  // Get keys from the first row, assuming all rows have similar structure
  const keys = Object.keys(data[0]);

  // Map keys to header configuration (customize labels and icons here)
  return keys.map(key => {
    let label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase()); // Basic formatting
    let icon: React.ElementType | undefined = undefined;

    // Example icon mapping (extend as needed)
    if (key.toLowerCase().includes('mail') || key.toLowerCase().includes('upn')) icon = Mail;
    if (key.toLowerCase().includes('sent')) icon = Send;
    if (key.toLowerCase().includes('received') || key.toLowerCase().includes('inbox')) icon = Inbox;

    // Specific overrides based on known keys from types like SharedMailboxTrafficStats
    if (key === 'sharedMailboxName') label = 'Shared Mailbox Name';
    if (key === 'sharedMailboxUPN') label = 'Shared Mailbox UPN';
    if (key === 'totalMailsSent') label = 'Total Mails Sent';
    if (key === 'totalMailsReceived') label = 'Total Mails Received';
    if (key === 'externalMailsSent') label = 'External Mails Sent';
    if (key === 'externalMailsReceived') label = 'External Mails Received';
     if (key === 'SkuId') label = 'SKU ID';
     if (key === 'SkuPartNumber') label = 'License Name';
     if (key === 'ConsumedUnits') label = 'Assigned';
     if (key === 'TotalUnits') label = 'Total';
     // Add more overrides for other report types


    return { key, label, icon };
  });
};

// Function to convert data to CSV string
const convertToCSV = (data: TableDataRow[], headers: { key: string; label: string }[]): string => {
  if (!data || data.length === 0) return '';

  const headerRow = headers.map(h => `"${h.label.replace(/"/g, '""')}"`).join(',');
  const dataRows = data.map(row =>
    headers
      .map(header => {
        const value = row[header.key];
        const stringValue = value === null || value === undefined ? '' : String(value);
        return `"${stringValue.replace(/"/g, '""')}"`; // Escape double quotes
      })
      .join(',')
  );

  return [headerRow, ...dataRows].join('\n');
};


export function ReportTable({ data, caption = "Report Data" }: ReportTableProps) {
  const [filter, setFilter] = useState('');
  const [sortKey, setSortKey] = useState<SortKey<TableDataRow>>(null);
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  // Generate headers dynamically based on the data
  const headers = useMemo(() => generateHeaders(data), [data]);

  const filteredData = useMemo(() => {
    if (!filter) return data;
    const lowerCaseFilter = filter.toLowerCase();
    return data.filter((row) =>
      headers.some(({ key }) => // Filter based on displayed columns only
         row[key] !== null &&
         row[key] !== undefined &&
         String(row[key]).toLowerCase().includes(lowerCaseFilter)
      )
    );
  }, [data, filter, headers]);

  const sortedData = useMemo(() => {
    if (!sortKey || !headers.find(h => h.key === sortKey)) return filteredData; // Ensure sortKey is valid

    return [...filteredData].sort((a, b) => {
      const aValue = a[sortKey];
      const bValue = b[sortKey];

      // Handle null or undefined values by pushing them to the end
       if (aValue == null && bValue == null) return 0;
       if (aValue == null) return sortDirection === 'asc' ? 1 : -1;
       if (bValue == null) return sortDirection === 'asc' ? -1 : 1;


      // Basic comparison for numbers and strings
       if (typeof aValue === 'number' && typeof bValue === 'number') {
        return sortDirection === 'asc' ? aValue - bValue : bValue - aValue;
      }
       if (typeof aValue === 'string' && typeof bValue === 'string') {
         // Case-insensitive sort for strings
         const comparison = aValue.toLowerCase().localeCompare(bValue.toLowerCase());
         return sortDirection === 'asc' ? comparison : -comparison;
       }

      // Fallback for other types (basic string conversion sort)
       const stringA = String(aValue).toLowerCase();
       const stringB = String(bValue).toLowerCase();
       const comparison = stringA.localeCompare(stringB);
       return sortDirection === 'asc' ? comparison : -comparison;
    });
  }, [filteredData, sortKey, sortDirection, headers]);

  const handleSort = (key: SortKey<TableDataRow>) => {
    if (sortKey === key) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortKey(key);
      setSortDirection('asc');
    }
  };

   const renderValue = (value: any) => {
    if (value === null || value === undefined) {
      return <span className="text-muted-foreground">-</span>;
    }
    if (typeof value === 'boolean') {
        return value ? 'Yes' : 'No';
    }
    if (value instanceof Date) {
        return value.toLocaleString(); // Or format as needed
    }
    // Simple array/object rendering (can be expanded)
     if (Array.isArray(value)) {
       return value.join(', ');
     }
     if (typeof value === 'object') {
       return JSON.stringify(value); // Basic fallback
     }

    // Add 'Mail(s)' suffix for specific known count columns (example)
    // This part might need refinement based on actual data types/keys
    // if (typeof value === 'number' && sortKey && String(sortKey).toLowerCase().includes('mail')) {
    //   return `${value} Mail${value !== 1 ? 's' : ''}`;
    // }
    return String(value);
  };


  // Download functionality using file-saver
  const handleDownload = useCallback(() => {
    if (!sortedData || sortedData.length === 0) {
      alert("No data available to download.");
      return;
    }

    try {
      const csvData = convertToCSV(sortedData, headers);
      const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
      const fileName = `${caption?.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'report'}_${new Date().toISOString().split('T')[0]}.csv`;
      saveAs(blob, fileName);
    } catch (error) {
        console.error("Error generating CSV:", error);
        alert("Failed to generate CSV file.");
    }
  }, [sortedData, headers, caption]);


  return (
    <div className="w-full">
      <div className="flex items-center py-4 gap-2">
        <Input
          placeholder="Filter results..."
          value={filter}
          onChange={(event) => setFilter(event.target.value)}
          className="max-w-sm"
          aria-label="Filter table data"
        />
         <Button
          variant="outline"
          size="sm"
          onClick={handleDownload}
          aria-label="Download report data as CSV"
          className="ml-auto" // Pushes the button to the right
          disabled={!sortedData || sortedData.length === 0} // Disable if no data
        >
          <Download className="mr-2 h-4 w-4" />
          Export CSV
        </Button>
      </div>
      <div className="rounded-md border overflow-auto max-h-[calc(100vh-280px)]"> {/* Adjusted max-h */}
        <Table>
          <TableCaption>{caption} ({sortedData.length} rows)</TableCaption>
          <TableHeader className="sticky top-0 bg-card z-10">
            <TableRow>
              {headers.map((header) => (
                <TableHead key={header.key} className="whitespace-nowrap">
                  <Button
                    variant="ghost"
                    onClick={() => handleSort(header.key)}
                    className="px-2 py-1 h-auto -ml-2"
                    aria-label={`Sort by ${header.label}`}
                  >
                    {header.icon && <header.icon className="mr-2 h-4 w-4 text-muted-foreground shrink-0" />}
                    <span className="truncate">{header.label}</span>
                    {sortKey === header.key && (
                      <ArrowUpDown className="ml-2 h-4 w-4 shrink-0" />
                    )}
                  </Button>
                </TableHead>
              ))}
            </TableRow>
          </TableHeader>
          <TableBody>
            {sortedData.length > 0 ? (
              sortedData.map((row, index) => (
                <TableRow key={index} data-testid={`report-row-${index}`}>
                 {headers.map((header) => (
                    <TableCell key={`${index}-${header.key}`} className="whitespace-nowrap">
                       {renderValue(row[header.key])}
                    </TableCell>
                  ))}
                </TableRow>
              ))
            ) : (
              <TableRow>
                <TableCell colSpan={headers.length} className="h-24 text-center">
                  {filter ? 'No results matching your filter.' : 'No data available.'}
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
