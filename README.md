# BBG_V3COM_ExcelDownloader

A VBA macro to download Bloomberg data via the V3COM API. Terminal subscription is required.

INSTRUCTIONS
1. Replace the placeholder Bloomberg tickers in Column A of the "bb_data" worksheet with the required tickers, to which you need to download the data for.
2. For the set of placeholder tickers, Column B holds the country code of the ticker while Column C provides the description of the ticker. Data that is downloaded will be saved in a csv file with the following concatenated naming convention:
  Column B + "_" + Column C + ".csv"
  Replace the values in these columns according to your preferred file naming convention.
3. The macro is set to download the last price ("PX_LAST") of the tickers listed in Column A described in Point 1. See section "CUSTOM CONFIGURATION" to change the type of data to be downloaded.

FURTHER CONFIGURATIONS
1. Data History Period - Set request parameters in the MakeRequest sub-routine in the Data Control class module.
2. Download Folder - Set the download folder for the csv files in the create_CsvFile sub-routine in the Data Control class module.
3. Download Data Field - Set to "PX_LAST" by default. Change to a desired data field in the DownloadData sub-routine in the Module1 module.
