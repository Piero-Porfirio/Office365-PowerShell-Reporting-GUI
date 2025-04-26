
'use client';

import React, { useState, useMemo } from 'react';
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
import type { SharedMailboxTrafficStats } from '@/types/reporting'; // Adjust type import as needed

type SortKey = keyof SharedMailboxTrafficStats | null;
type SortDirection = 'asc' | 'desc';

interface ReportTableProps {
  data: SharedMailboxTrafficStats[]; // Make data type more specific
  caption?: string;
}

export function ReportTable({ data, caption = "Report Data" }: ReportTableProps) {
  const [filter, setFilter] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>(null);
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  const headers: { key: keyof SharedMailboxTrafficStats; label: string; icon?: React.ElementType }[] = [
    { key: 'month', label: 'Month' },
    { key: 'sharedMailboxName', label: 'Shared Mailbox Name' },
    { key: 'sharedMailboxUPN', label: 'Shared Mailbox UPN', icon: Mail },
    { key: 'totalMailsSent', label: 'Total Mails Sent', icon: Send },
    { key: 'totalMailsReceived', label: 'Total Mails Received', icon: Inbox },
    { key: 'externalMailsSent', label: 'External Mails Sent', icon: Send },
    { key: 'externalMailsReceived', label: 'External Mails Received', icon: Inbox },
  ];

  const filteredData = useMemo(() => {
    if (!filter) return data;
    const lowerCaseFilter = filter.toLowerCase();
    return data.filter((row) =>
      Object.values(row).some(
        (value) =>
          value !== null &&
          value !== undefined &&
          value.toString().toLowerCase().includes(lowerCaseFilter)
      )
    );
  }, [data, filter]);

  const sortedData = useMemo(() => {
    if (!sortKey) return filteredData;

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
         const comparison = aValue.localeCompare(bValue);
         return sortDirection === 'asc' ? comparison : -comparison;
       }

      // Fallback for other types (less reliable sorting)
      if (aValue < bValue) return sortDirection === 'asc' ? -1 : 1;
      if (aValue > bValue) return sortDirection === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredData, sortKey, sortDirection]);

  const handleSort = (key: SortKey) => {
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
    // Add 'Mail(s)' suffix for count columns
    if (typeof value === 'number') {
      return `${value} Mail${value !== 1 ? 's' : ''}`;
    }
    return value.toString();
  };


  // Placeholder for download functionality
  const handleDownload = () => {
    console.log("Download requested:", sortedData);
    // Implement actual CSV/Excel download logic here
    alert("Download functionality not yet implemented.");
  };


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
          aria-label="Download report data"
          className="ml-auto" // Pushes the button to the right
        >
          <Download className="mr-2 h-4 w-4" />
          Download
        </Button>
      </div>
      <div className="rounded-md border overflow-auto max-h-[calc(100vh-250px)]"> {/* Adjust max-h as needed */}
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
                    {header.icon && <header.icon className="mr-2 h-4 w-4 text-muted-foreground" />}
                    {header.label}
                    {sortKey === header.key && (
                      <ArrowUpDown className="ml-2 h-4 w-4" />
                    )}
                  </Button>
                </TableHead>
              ))}
            </TableRow>
          </TableHeader>
          <TableBody>
            {sortedData.length > 0 ? (
              sortedData.map((row, index) => (
                <TableRow key={index}>
                 {headers.map((header) => (
                    <TableCell key={header.key} className="whitespace-nowrap">
                       {renderValue(row[header.key])}
                    </TableCell>
                  ))}
                </TableRow>
              ))
            ) : (
              <TableRow>
                <TableCell colSpan={headers.length} className="h-24 text-center">
                  No results found.
                </TableCell>
              </TableRow>
            )}
          </TableBody>
            {/* Optional Footer */}
          {/* <TableFooter>
            <TableRow>
              <TableCell colSpan={headers.length}>Total Rows: {sortedData.length}</TableCell>
            </TableRow>
          </TableFooter> */}
        </Table>
      </div>
    </div>
  );
}
