
'use client';

import React, { useMemo } from 'react';
import { Pie, PieChart, ResponsiveContainer, Cell, Tooltip, Legend } from 'recharts';
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
  CardDescription,
} from '@/components/ui/card';
import {
  ChartContainer,
  ChartTooltipContent,
  ChartLegend,
  ChartLegendContent,
  ChartStyle
} from "@/components/ui/chart";
import type { LicenseUtilizationData } from '@/types/reporting';
import { cn } from '@/lib/utils';

interface ReportComponentProps {
  data: any; // Expect generic data here
  caption?: string;
}

// Enhance data type expected by this specific chart
type ChartLicenseData = LicenseUtilizationData & { AvailableUnits?: number };

// Define colors for the chart segments
const COLORS = ['hsl(var(--chart-1))', 'hsl(var(--chart-3))', 'hsl(var(--chart-5))', 'hsl(var(--chart-2))', 'hsl(var(--chart-4))'];

export function LicenseUtilizationChart({ data, caption = "License Utilization" }: ReportComponentProps) {

  // Ensure data is treated as LicenseUtilizationData[] for processing
  const licenseData = useMemo(() => (data as LicenseUtilizationData[] || []), [data]);

  // Calculate available units and aggregate data if multiple entries for the same license type exist
  const aggregatedData = useMemo(() => {
    const map = new Map<string, { SkuPartNumber: string; ConsumedUnits: number; TotalUnits: number; AvailableUnits: number }>();

    licenseData.forEach(item => {
        const key = item.SkuPartNumber;
        const existing = map.get(key);
        const consumed = item.ConsumedUnits || 0;
        const total = item.TotalUnits || 0;
        const available = Math.max(0, total - consumed); // Ensure available is not negative

        if (existing) {
            existing.ConsumedUnits += consumed;
            existing.TotalUnits += total;
            existing.AvailableUnits += available;
        } else {
            map.set(key, {
                SkuPartNumber: item.SkuPartNumber,
                ConsumedUnits: consumed,
                TotalUnits: total,
                AvailableUnits: available,
            });
        }
    });

    return Array.from(map.values());
  }, [licenseData]);

  // Prepare data specifically for the Pie chart (Assigned vs Available per license type)
  const chartData = useMemo(() => {
    return aggregatedData.map(item => ({
      name: item.SkuPartNumber,
      assigned: item.ConsumedUnits,
      available: item.AvailableUnits,
      total: item.TotalUnits,
    }));
  }, [aggregatedData]);

  const chartConfig = useMemo(() => {
      const config: any = {};
      chartData.forEach((item, index) => {
          config[item.name] = {
              label: item.name,
              // Assign colors cyclically
              color: COLORS[index % COLORS.length],
          };
      });
       // Add config for 'Assigned' and 'Available' specifically if needed for tooltips/legends
       config.assigned = { label: "Assigned", color: 'hsl(var(--chart-1))' }; // Example, might need adjustment
       config.available = { label: "Available", color: 'hsl(var(--chart-3))' }; // Example, might need adjustment
      return config;
  }, [chartData]);

  if (!chartData || chartData.length === 0) {
    return <p>No data available for chart.</p>;
  }

  // Calculate overall totals for the main chart description
  const totalConsumed = aggregatedData.reduce((sum, item) => sum + item.ConsumedUnits, 0);
  const totalAvailable = aggregatedData.reduce((sum, item) => sum + item.AvailableUnits, 0);
  const totalOverall = totalConsumed + totalAvailable;
  const utilizationPercentage = totalOverall > 0 ? ((totalConsumed / totalOverall) * 100).toFixed(1) : 0;

  return (
    <ChartContainer config={chartConfig} className="mx-auto aspect-square max-h-[400px] w-full">
        <ResponsiveContainer width="100%" height={350}>
        <PieChart>
          <ChartStyle id="license-pie" config={chartConfig} />
          <Tooltip
            cursor={false}
            content={<ChartTooltipContent hideLabel nameKey="name" />}
          />
          <Pie
            data={chartData}
            dataKey="assigned" // Primary value for the pie segments
            nameKey="name"
            cx="50%"
            cy="50%"
            outerRadius={100}
            innerRadius={60} // Make it a Donut chart
            fill="hsl(var(--chart-1))" // Base fill, overridden by Cell
            strokeWidth={2}
            labelLine={false}
            // label={({ cx, cy, midAngle, innerRadius, outerRadius, percent, index, name }) => {
            //   const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
            //   const x = cx + radius * Math.cos(-midAngle * (Math.PI / 180));
            //   const y = cy + radius * Math.sin(-midAngle * (Math.PI / 180));
            //   // Only show label if percentage is significant
            //   return percent > 0.05 ? (
            //     <text x={x} y={y} fill="hsl(var(--foreground))" textAnchor={x > cx ? 'start' : 'end'} dominantBaseline="central" className="text-xs fill-foreground">
            //       {`${name} (${(percent * 100).toFixed(0)}%)`}
            //     </text>
            //   ) : null;
            // }}
          >
            {chartData.map((entry, index) => (
              <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
            ))}
          </Pie>
          <Legend content={<ChartLegendContent nameKey="name" />} />
          {/* Optional: Add text in the center of the Donut */}
          <text
            x="50%"
            y="50%"
            textAnchor="middle"
            dominantBaseline="middle"
            className="fill-foreground text-lg font-semibold"
          >
            {`${utilizationPercentage}%`}
          </text>
          <text
            x="50%"
            y="50%"
            textAnchor="middle"
            dominantBaseline="middle"
            dy="1.2em" // Position below the percentage
            className="fill-muted-foreground text-xs"
          >
            Used
          </text>
        </PieChart>
        </ResponsiveContainer>
    </ChartContainer>
  );
}
