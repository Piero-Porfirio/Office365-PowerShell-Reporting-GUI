
'use client';

import React, { useMemo } from 'react';
import { Bar, BarChart, CartesianGrid, XAxis, YAxis, ResponsiveContainer, Tooltip, Legend } from 'recharts';
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
  CardDescription,
} from '@/components/ui/card';
import {
  ChartContainer,
  ChartTooltip,
  ChartTooltipContent,
  ChartLegend,
  ChartLegendContent,
  ChartStyle
} from "@/components/ui/chart";
import type { SharedMailboxTrafficStats } from '@/types/reporting';
import { cn } from '@/lib/utils';

interface SharedMailboxTrafficChartProps {
  data: SharedMailboxTrafficStats[];
  caption?: string;
}

export function SharedMailboxTrafficChart({ data, caption = "Shared Mailbox Traffic" }: SharedMailboxTrafficChartProps) {

  // Aggregate data per mailbox across different months if necessary, or prepare for chart
  const aggregatedData = useMemo(() => {
    const map = new Map<string, {
        sharedMailboxName: string;
        totalMailsSent: number;
        totalMailsReceived: number;
        externalMailsSent: number;
        externalMailsReceived: number;
    }>();

    data.forEach(item => {
        const key = item.sharedMailboxUPN; // Use UPN as unique key
        const existing = map.get(key);

        // Ensure nulls are treated as 0
        const totalSent = item.totalMailsSent ?? 0;
        const totalReceived = item.totalMailsReceived ?? 0;
        const externalSent = item.externalMailsSent ?? 0;
        const externalReceived = item.externalMailsReceived ?? 0;

        if (existing) {
            existing.totalMailsSent += totalSent;
            existing.totalMailsReceived += totalReceived;
            existing.externalMailsSent += externalSent;
            existing.externalMailsReceived += externalReceived;
        } else {
            map.set(key, {
                sharedMailboxName: item.sharedMailboxName,
                totalMailsSent: totalSent,
                totalMailsReceived: totalReceived,
                externalMailsSent: externalSent,
                externalMailsReceived: externalReceived,
            });
        }
    });
     // Sort by total received emails, descending
     return Array.from(map.values()).sort((a, b) => b.totalMailsReceived - a.totalMailsReceived);
  }, [data]);

  // Define chart configuration for colors and labels
  const chartConfig = useMemo(() => ({
    totalMailsReceived: {
      label: "Received",
      color: "hsl(var(--chart-1))", // Blue
    },
    totalMailsSent: {
      label: "Sent",
      color: "hsl(var(--chart-2))", // Green
    },
    sharedMailboxName: {
        label: "Mailbox",
    },
  }), []);

  if (!aggregatedData || aggregatedData.length === 0) {
    return <p>No data available for chart.</p>;
  }

  return (
      <ChartContainer config={chartConfig} className="min-h-[300px] w-full">
           <ResponsiveContainer width="100%" height={350}>
                <BarChart data={aggregatedData} layout="vertical" margin={{ left: 20, right: 30 }}>
                    <ChartStyle id="mailbox-traffic-bar" config={chartConfig} />
                    <CartesianGrid horizontal={false} strokeDasharray="3 3" />
                    <XAxis type="number" dataKey="totalMailsReceived" hide />
                    <YAxis
                        dataKey="sharedMailboxName"
                        type="category"
                        tickLine={false}
                        tickMargin={10}
                        axisLine={false}
                        // Use tick formatter for potentially long names
                        tickFormatter={(value) => value.length > 20 ? `${value.substring(0, 18)}...` : value}
                        width={150} // Adjust width based on expected label length
                    />
                    <Tooltip
                        cursor={false}
                        content={<ChartTooltipContent indicator="line" />}
                    />
                    <Legend content={<ChartLegendContent />} />
                    <Bar dataKey="totalMailsReceived" fill="var(--color-totalMailsReceived)" radius={4} name="Received" />
                    <Bar dataKey="totalMailsSent" fill="var(--color-totalMailsSent)" radius={4} name="Sent" />
                </BarChart>
           </ResponsiveContainer>
      </ChartContainer>
  );
}
