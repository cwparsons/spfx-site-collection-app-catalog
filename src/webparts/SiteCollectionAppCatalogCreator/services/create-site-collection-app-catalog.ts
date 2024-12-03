import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

import escapeXml from "xml-escape";

type AppCatalogInfo = {
  SchemaVersion: string;
  LibraryVersion: string;
  ErrorInfo: ErrorInfo;
  TraceCorrelationId: string;
};

type ErrorInfo = {
  ErrorMessage: string;
  ErrorValue: unknown;
  TraceCorrelationId: string;
  ErrorCode: number;
  ErrorTypeName: string;
};

export const createCollectionSiteCollectionAppCatalog = async (
  context: WebPartContext,
  url: string
): Promise<AppCatalogInfo[]> => {
  const contextinfoRequest = await context.spHttpClient.post(
    `${context.pageContext.site.absoluteUrl}/_api/contextinfo`,
    SPHttpClient.configurations.v1,
    {}
  );
  const contextinfo = await contextinfoRequest.json();

  const processQueryRequest = await context.spHttpClient.post(
    `${context.pageContext.site.absoluteUrl}/_vti_bin/client.svc/ProcessQuery`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        "X-RequestDigest": contextinfo.FormDigestValue,
      },
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="GO Development Tools" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${escapeXml(
        url
      )}</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">${escapeXml(
        url
      )}</Parameter></Parameters></Method></ObjectPaths></Request>`,
    }
  );

  return processQueryRequest.json();
};
