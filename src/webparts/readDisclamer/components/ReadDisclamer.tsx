import * as React from 'react';
import { useState, useEffect } from 'react';
import type { IReadDisclamerProps } from './IReadDisclamerProps';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  Checkbox,
  Text,
  IStackTokens,
  ITheme,
  Stack
} from "office-ui-fabric-react";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Placeholder } from "@pnp/spfx-controls-react"; ///lib/Placeholder

const ReadDisclamer: React.FunctionComponent<IReadDisclamerProps> = (
  props
) => {
  const [showMessage, setShowMessage] = useState<boolean>(true);

  const themeVariantColors: IReadonlyTheme | undefined = props.themeVariant;

  const fetchData = async (): Promise<void> => {
    // Another approach to get current user info could be
    // e.g this.context.pageContext.legacyPageContext.userId

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = await sp.web.lists
      .getById(props.storageList).items
      //.filter(`Author/EMail eq '${encodeURIComponent(props.currentUser?.Email || '')}' and Title eq '${props.documentTitle}'`)
      .filter(`Author eq '${props.currentUser?.Id}' and Title eq '${props.documentTitle}'`)
      .top(1)
      .select("Title,Author/ID,Author/Name,Author/Title,Author/EMail").expand("Author")  // fields to return
      .get();

    setShowMessage(items.length === 0)
  };

  useEffect(() => {
    const fetchDataAsync = async (): Promise<void> => {
      if (props.storageList && props.storageList !== "") {
        await fetchData();
      }
    };

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    fetchDataAsync();
  }, [props]);

  const _onConfigure = (): void => {
    // Context of the web part
    props.context.propertyPane.open();
  };

  const _onChange = async (ev: React.FormEvent<HTMLElement>, isChecked: boolean): Promise<void> => {
    await sp.web.lists.getById(props.storageList).items.add({
      Title: props.documentTitle
    });
    setShowMessage(false);
  }

  const mainStackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10,
  };

  return props.configured ? (
    <Stack style={{ backgroundColor: themeVariantColors?.semanticColors?.bodyBackground }}>
      {showMessage ? (
        <Stack
          style={{ color: themeVariantColors?.semanticColors?.bodyText }}
          tokens={mainStackTokens}
        >
          <Text>{props.acknowledgementMessage}</Text>
          <Text variant="large">{props.documentTitle}</Text>
          <Checkbox
            theme={props.themeVariant as ITheme}
            label={props.acknowledgementMessage}
            onChange={_onChange}
          />
        </Stack>
      ) : (
        <Stack style={{ color: themeVariantColors?.semanticColors?.bodyText }}>
          <Text variant="large">{props.documentTitle}</Text>
          <Text>{props.readMessage}</Text>
        </Stack>
      )}
    </Stack>
  ) : (
    <Placeholder
      iconName="Edit"
      iconText="Configure Read Disclamer"
      description="Please configure the web part by choosing a list."
      buttonLabel="Configure"
      onConfigure={_onConfigure}
    />
  );
}

export default ReadDisclamer;
