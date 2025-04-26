
'use client';

import React from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Separator } from '@/components/ui/separator';
import { Globe, Key, Building } from 'lucide-react';
import { useToast } from '@/hooks/use-toast';

export default function SettingsPage() {
  const { toast } = useToast();

  const handleConnectAzure = () => {
    // Placeholder logic for connecting to Azure
    console.log('Attempting to connect to Azure...');
    // In a real app, this would trigger an OAuth flow or prompt for credentials/secrets
    toast({
      title: 'Azure Connection (Placeholder)',
      description: 'Azure connection functionality not yet implemented.',
    });
  };

  const handleConnectOffice365 = () => {
    // Placeholder logic for connecting to Office 365
    console.log('Attempting to connect to Office 365...');
    // Similar to Azure, this would involve authentication
    toast({
      title: 'Office 365 Connection (Placeholder)',
      description: 'Office 365 connection functionality not yet implemented.',
    });
  };


  return (
    <main className="flex flex-col items-center p-4 md:p-8 lg:p-12 bg-background min-h-screen">
      <div className="w-full max-w-4xl">
        <h1 className="text-3xl font-bold mb-6 text-center text-primary">
          Settings
        </h1>

        <Card className="w-full shadow-lg border-border/50">
          <CardHeader>
            <CardTitle className="text-xl">Tenant Connections</CardTitle>
            <CardDescription>
              Connect your Azure and Office 365 tenants to enable PowerShell execution and reporting.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {/* Azure Connection Section */}
            <div className="space-y-4 p-4 border rounded-md border-input">
              <h3 className="text-lg font-semibold flex items-center">
                <Globe className="mr-2 h-5 w-5 text-blue-500" /> Azure Tenant
              </h3>
              <p className="text-sm text-muted-foreground">
                Configure connection details for your Azure subscription. This typically involves setting up an App Registration with appropriate permissions.
              </p>
              {/* Placeholder for connection status/details */}
              <div className="text-sm text-muted-foreground italic">
                Status: Not Connected
              </div>
              {/* <div className="space-y-2">
                <Label htmlFor="azureTenantId">Tenant ID</Label>
                <Input id="azureTenantId" placeholder="Enter Azure Tenant ID" disabled />
              </div>
              <div className="space-y-2">
                <Label htmlFor="azureClientId">Client ID</Label>
                <Input id="azureClientId" placeholder="Enter App Registration Client ID" disabled />
              </div>
              <div className="space-y-2">
                <Label htmlFor="azureClientSecret">Client Secret</Label>
                <Input id="azureClientSecret" type="password" placeholder="Enter Client Secret" disabled />
              </div> */}
              <Button onClick={handleConnectAzure} className="w-full md:w-auto">
                 Connect Azure (Not Implemented)
              </Button>
            </div>

            <Separator />

            {/* Office 365 Connection Section */}
            <div className="space-y-4 p-4 border rounded-md border-input">
               <h3 className="text-lg font-semibold flex items-center">
                 <Building className="mr-2 h-5 w-5 text-orange-500" /> Office 365 Tenant
              </h3>
               <p className="text-sm text-muted-foreground">
                Configure connection details for your Office 365 environment. Similar permissions via an App Registration are often required.
              </p>
               {/* Placeholder for connection status/details */}
              <div className="text-sm text-muted-foreground italic">
                 Status: Not Connected
               </div>
              {/* <div className="space-y-2">
                 <Label htmlFor="o365TenantId">Tenant ID</Label>
                 <Input id="o365TenantId" placeholder="Enter Office 365 Tenant ID" disabled />
               </div>
               <div className="space-y-2">
                 <Label htmlFor="o365ClientId">Client ID</Label>
                 <Input id="o365ClientId" placeholder="Enter App Registration Client ID" disabled />
               </div>
               <div className="space-y-2">
                 <Label htmlFor="o365ClientSecret">Client Secret</Label>
                 <Input id="o365ClientSecret" type="password" placeholder="Enter Client Secret" disabled />
               </div> */}
               <Button onClick={handleConnectOffice365} className="w-full md:w-auto">
                 Connect Office 365 (Not Implemented)
               </Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </main>
  );
}
