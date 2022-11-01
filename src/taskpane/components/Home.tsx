import React, { useEffect, useState } from "react";
import { DefaultButton, Label, Stack, TextField } from "@fluentui/react";
/* global console, Excel  */
const Home = () => {
  const [currentSelection, setCurrentSelection] = useState("");
  const [fieldName, setFieldName] = useState("");
  useEffect(() => {
    const registerOnChangedEvent = async () => {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items");
        await context.sync();

        worksheets.items.forEach(async (sheet) => {
          sheet.onSelectionChanged.add(handleSelectionChange);
          await context.sync();
        });
      });
    };
    const handleSelectionChange = async (event: Excel.WorksheetSelectionChangedEventArgs) => {
      await Excel.run(async (context) => {
        const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
        activeWorksheet.load("name");
        await context.sync();
        setCurrentSelection(`${activeWorksheet.name}!${event.address}`);
        console.log("Address of current selection: " + activeWorksheet.name + "!" + event.address);
      });
    };
    registerOnChangedEvent();
  }, []);
  useEffect(() => {
    const getValue = async (currentSelection) => {
      const property = await getCustomProperties(currentSelection);

      if (property != undefined) {
        setFieldName(property.value);
      } else {
        setFieldName("");
      }
    };
    getValue(currentSelection);
  }, [currentSelection]);
  const setCustomProperties = async () => {
    try {
      await Excel.run(async (context) => {
        const properties = context.workbook.properties.custom;
        properties.add(currentSelection, fieldName);
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };
  const getCustomProperties = async (key: string) => {
    try {
      return await Excel.run(async (context) => {
        const properties = context.workbook.properties.custom;
        properties.load(["key", "value"]);
        await context.sync();
        const result = properties.items.filter((property) => property.key == key);
        if (result.length > 0) return result[0];
        return undefined;
      });
    } catch (error) {
      console.error(error);
      return undefined;
    }
  };
  return (
    <>
      <Stack verticalFill padding={10}>
        <Label>Current Selection: {currentSelection}</Label>
        <TextField
          label="Field Name"
          placeholder="Please enter field name"
          value={fieldName}
          onChange={(_event, newValue) => {
            setFieldName(newValue);
          }}
        />
        <DefaultButton onClick={setCustomProperties}>Save</DefaultButton>
      </Stack>
    </>
  );
};
export default Home;
