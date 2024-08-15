import { spfi, SPFI, SPFx } from "@pnp/sp";
import { SPHttpClient } from "@microsoft/sp-http";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CellRender } from "../common/ColumnDetails";
import { getColumnMaxWidth, getColumnMinWidth } from "../utils/columnUtils";
import { ConstructedFilter } from "./PanelComponent";

export interface ITerm {
  Id: string;
  Name: string;
  parentId: string;
  Children?: ITerm[];
}

export interface TermSet {
  setId: string;
  terms: ITerm[];
}

export interface FilterDetail {
  filterColumn: string;
  filterColumnType: string;
  values: string[];
}

export interface IColumnInfo {
  InternalName: string;
  DisplayName: string;
  MinWidth: number;
  ColumnType: string;
  MaxWidth: number;
  OnRender?: (items: any) => JSX.Element;
}
class PagesService {
  private _sp: SPFI;

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }

  /**
   * Fetch distinct values for a given column from a list of items.
   * @param {string} columnName - The name of the column to fetch distinct values for.
   * @param {any[]} values - The list of items to extract distinct values from.
   * @returns {Promise<string[] | ConstructedFilter[]>} - A promise that resolves to an array of distinct values.
   */
  getDistinctValues = async (
    columnName: string,
    columnType: string,
    values: any
  ): Promise<(string | ConstructedFilter)[]> => {
    try {
      const items = values; // The list of items to fetch distinct values from.

      // Extract distinct values from the column
      const distinctValues: (string | ConstructedFilter)[] = [];
      const seenValues = new Set<string | ConstructedFilter>(); // A set to keep track of seen values to avoid duplicates.

      items.forEach((item: any) => {
        switch (columnType) {
          case "TaxonomyFieldTypeMulti":
            if (item[columnName] && item[columnName].length > 0) {
              // Extract distinct values from the column
              item[columnName].forEach((category: any) => {
                const uniqueValue = category.Label;
                if (!seenValues.has(uniqueValue)) {
                  seenValues.add(uniqueValue);
                  distinctValues.push(uniqueValue);
                }
              });
            }
            break;
          case "DateTime":
            let uniqueDateValue = item[columnName]; // The value of the column for the current item.
            // Handle ISO date strings by extracting only the date part
            uniqueDateValue = new Date(uniqueDateValue)
              .toISOString()
              .split("T")[0];

            if (!seenValues.has(uniqueDateValue)) {
              seenValues.add(uniqueDateValue);
              distinctValues.push(uniqueDateValue);
            }
            break;
          case "User":
            const userValue = item[columnName];
            if (
              userValue &&
              userValue.Title &&
              !seenValues.has(userValue.Title)
            ) {
              seenValues.add(userValue.Title);
              const user: ConstructedFilter = {
                text: userValue.Title,
                value: userValue.Id,
              };
              distinctValues.push(user);
            }
            break;
          case "Number":
            const uniqueNumberValue = item[columnName]; // The value of the column for the current item.

            if (!seenValues.has(uniqueNumberValue)) {
              seenValues.add(uniqueNumberValue);
              distinctValues.push(uniqueNumberValue);
            }
            break;
          case "Choice":
            const uniqueChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueChoiceValue) {
              if (!seenValues.has(uniqueChoiceValue)) {
                seenValues.add(uniqueChoiceValue);
                distinctValues.push(uniqueChoiceValue);
              }
            }
            break;
          case "URL":
            const uniqueUrlChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueUrlChoiceValue && uniqueUrlChoiceValue.Url) {
              if (!seenValues.has(uniqueUrlChoiceValue.Url)) {
                seenValues.add(uniqueUrlChoiceValue.Url);
                distinctValues.push(uniqueUrlChoiceValue.Url);
              }
            }
            break;
          case "Computed":
            const uniqueCompChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueCompChoiceValue) {
              if (!seenValues.has(uniqueCompChoiceValue.split(".")[0])) {
                seenValues.add(uniqueCompChoiceValue.split(".")[0]);
                distinctValues.push(uniqueCompChoiceValue.split(".")[0]);
              }
            }
            break;
          default:
            const uniqueValue = item[columnName]; // The value of the column for the current item.
            if (uniqueValue) {
              if (!seenValues.has(uniqueValue)) {
                seenValues.add(uniqueValue);
                distinctValues.push(uniqueValue);
              }
            }
            break;
        }
      });

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  /**
   * Retrieves a page of filtered Site Pages items.
   *
   * @param viewId The selected view id
   * @param pageNumber The page number to retrieve (1-indexed).
   * @param pageSize The number of items to retrieve per page. Defaults to 10.
   * @param orderBy The column to sort the items by. Defaults to "Created".
   * @param isAscending Whether to sort in ascending or descending order. Defaults to true.
   * @param folderPath The folder path to search in. Defaults to "" (root of the site).
   * @param searchText Text to search for in the Title, Article ID, or Modified columns.
   * @param filters An array of FilterDetail objects to apply to the query.
   * @returns A promise that resolves with an array of items.
   */
  getFilteredPages = async (
    pageNumber: number,
    pageSize: number = 10,
    orderBy: string = "Created",
    isAscending: boolean = true,
    folderPath: string = "",
    searchText: string = "",
    filters: FilterDetail[],
    columnInfos: IColumnInfo[]
  ) => {
    try {
      const skip = (pageNumber - 1) * pageSize;
      const list = this._sp.web.lists.getByTitle("Site Pages");

      // Default columns to always include
      const allFieldsSet = new Set<string>();
      const expandFieldsSet = new Set<string>();

      columnInfos.forEach((col) => {
        if (col.ColumnType === "User") {
          allFieldsSet.add(`${col.InternalName}/Id`);
          allFieldsSet.add(`${col.InternalName}/Title`);
          expandFieldsSet.add(col.InternalName.split("/")[0]);
        } else if (
          col.ColumnType === "TaxonomyFieldTypeMulti" ||
          col.ColumnType === "TaxonomyFieldType"
        ) {
          allFieldsSet.add(col.InternalName);
        } else {
          allFieldsSet.add(`${col.InternalName}`);
        }
      });

      const allFields: string[] = [
        "FileRef",
        "FileDirRef",
        "FSObjType",
        "Title",
        "Id",
      ];
      allFieldsSet.forEach((col) => allFields.push(col));

      const expandFields: string[] = [];
      expandFieldsSet.forEach((field) => expandFields.push(field));

      let filterQuery = `startswith(FileDirRef, '${folderPath}') and FSObjType eq 0${
        searchText
          ? ` and (substringof('${searchText}', Title) or Article_x0020_ID eq '${searchText}' or substringof('${searchText}', Modified))`
          : ""
      }`;

      filters.forEach((filter) => {
        if (filter.values.length > 0) {
          switch (filter.filterColumnType) {
            case "TaxonomyFieldTypeMulti":
              return;

            case "DateTime":
              const dateFilters = filter.values
                .map((value) => {
                  const startDate = new Date(value);
                  const endDate = new Date(value);
                  endDate.setDate(endDate.getDate() + 1);

                  return `${
                    filter.filterColumn
                  } ge datetime'${startDate.toISOString()}' and ${
                    filter.filterColumn
                  } lt datetime'${endDate.toISOString()}'`;
                })
                .join(" or ");
              if (dateFilters && dateFilters != "")
                filterQuery += ` and (${dateFilters})`;
              break;

            case "User":
              const userFilters = filter.values
                .map((value) => `${filter.filterColumn}/Id eq '${value}'`)
                .join(" or ");
              if (userFilters && userFilters != "")
                filterQuery += ` and (${userFilters})`;
              break;
            case "URL":
              const urlFilters = filter.values
                .map((value) => {
                  return `${filter.filterColumn}/Url eq '${value}'`;
                })
                .join(" or ");
              if (urlFilters && urlFilters != "")
                filterQuery += ` and (${urlFilters})`;
              break;

            case "Computed":
              if (
                filter.filterColumn === "Name" ||
                filter.filterColumn === "FileLeafRef" ||
                filter.filterColumn === "LinkFilename" ||
                filter.filterColumn === "LinkFilenameNoMenu"
              ) {
                const urlFilters = filter.values
                  .map((value) => {
                    return `Title eq '${value}'`;
                  })
                  .join(" or ");
                if (urlFilters && urlFilters != "")
                  filterQuery += ` and (${urlFilters})`;
              }
              break;

            default:
              const columnFilters = filter.values
                .map((value) => `${filter.filterColumn} eq '${value}'`)
                .join(" or ");
              if (columnFilters && columnFilters != "")
                filterQuery += ` and (${columnFilters})`;
              break;
          }
        }
      });

      const pages: any[] = await list.items
        .filter(filterQuery)
        .select(...allFields)
        .expand(...expandFields)
        .skip(skip)
        .top(5000)
        .orderBy(orderBy, isAscending)();

      const currentUser = await this._sp.web.currentUser();
      const userAlertsResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/alerts?$filter=UserId eq ${currentUser.Id} and substringof('Site Pages', Title)`,
        SPHttpClient.configurations.v1
      );
      const userAlerts = await userAlertsResponse.json();

      // Check subscription status for each page
      const subscribedPages: string[] = [];
      userAlerts.value.forEach((alert: any) => {
        // Extract the page name from the alert title
        const pageName = alert.Title.split("Site Pages: ")[1];
        subscribedPages.push(pageName);
      });

      // Add subscription status to each page item
      pages.forEach((page) => {
        const pageFileName = page.FileRef?.split("/").pop(); // Extract file name from FileRef
        page.Subscribed =
          subscribedPages.indexOf(pageFileName) !== -1 ? true : false;
      });

      return pages;
    } catch (error) {
      console.error("Error fetching filtered pages:", error);
      throw new Error("Error fetching filtered pages");
    }
  };
  /**
   * Retrieves the columns for a specified view in the SharePoint list.
   */
  public async getColumns(viewId: string): Promise<IColumnInfo[]> {
    const fields = await this._sp.web.lists
      .getByTitle("Site Pages")
      .views.getById(viewId)
      .fields();

    // Fetching detailed field information to get both internal and display names
    const fieldDetailsPromises = fields.Items.map((field: any) =>
      this._sp.web.lists
        .getByTitle("Site Pages")
        .fields.getByInternalNameOrTitle(field)()
    );

    const fieldDetails = await Promise.all(fieldDetailsPromises);

    return fieldDetails.map((field: any) => ({
      InternalName: field.InternalName,
      DisplayName: field.Title,
      ColumnType: field.TypeAsString,
      MinWidth: getColumnMinWidth(field.InternalName),
      MaxWidth: getColumnMaxWidth(field.InternalName),
      OnRender: (item: any) =>
        CellRender({
          columnName: field.InternalName,
          columnType: field.TypeAsString,
          item,
          context: this.context,
        }),
    }));
  }

  /**
   * Retrieves the details of a SharePoint list by its name.
   * @param {string} listName - The name of the list to retrieve details for.
   * @returns {Promise<any>} - A promise that resolves to the list details.
   */
  public async getListDetailsByName(listName: string): Promise<any> {
    try {
      const list = await this._sp.web.lists.getByTitle(listName)();
      return list;
    } catch (error) {
      console.error(`Error retrieving list details for ${listName}:`, error);
      throw new Error(`Error retrieving list details for ${listName}`);
    }
  }
}

export default PagesService;
