import * as React from "react";

import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { TextField } from "@fluentui/react/lib/TextField";
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { useState } from "react";

import { createCollectionSiteCollectionAppCatalog } from "../services/create-site-collection-app-catalog";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

export type SiteCollectionAppCatalogCreatorProps = {
  context: WebPartContext;
};

const SiteCollectionAppCatalogCreator = ({
  context,
}: SiteCollectionAppCatalogCreatorProps): React.ReactNode => {
  const [status, setStatus] = useState<
    "initial" | "loading" | "failure" | "success"
  >("initial");
  const [url, setUrl] = useState<string>();

  const onSubmit = async (): Promise<void> => {
    if (!url) return;

    setStatus("loading");

    try {
      const info = await createCollectionSiteCollectionAppCatalog(context, url);

      if (info[0].ErrorInfo) {
        setStatus("failure");
      } else {
        setStatus("success");
      }
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
    } catch (_e) {
      setStatus("failure");
    }
  };

  return (
    <>
      <WebPartTitle
        displayMode={DisplayMode.Read}
        title={"Site app catalog creator"}
        updateProperty={function (value: string): void {
          throw new Error("Function not implemented.");
        }}
      />

      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "start",
          gap: "1rem",
        }}
      >
        <TextField
          label="Absolute site URL"
          onChange={(_e, newValue) => {
            setUrl(newValue);
          }}
          placeholder="https://{tenant}.sharepoint.com/sites/{site}"
          type="url"
        />

        <PrimaryButton onClick={onSubmit}>Submit</PrimaryButton>

        {status === "loading" && (
          <Spinner label="Creating site app catalog..." />
        )}

        {status === "failure" && (
          <MessageBar messageBarType={MessageBarType.error}>
            An error occurred creating the site collection app catalog. It may
            <a
              href={`${url}/_layouts/15/viewlsts.aspx`}
              target="_blank"
              rel="noreferrer"
            >
              already be created
            </a>
            .
          </MessageBar>
        )}

        {status === "success" && (
          <MessageBar messageBarType={MessageBarType.success}>
            Created
            <a
              href={`${url}/_layouts/15/viewlsts.aspx`}
              target="_blank"
              rel="noreferrer"
            >
              site collection app catalog
            </a>
            .
          </MessageBar>
        )}
      </div>
    </>
  );
};

export default SiteCollectionAppCatalogCreator;
