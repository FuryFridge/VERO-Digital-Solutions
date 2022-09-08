Script won't work without some additional libraries that can be installed via [requirements.txt](requirements.txt)

1. Function `parse_args(args)` in the script checks the input parameters `-k/--keys` or `-c/--colored`:
   - If `-k` parameter is given, than it will show additional columns in the output file, else it will show only primary columns. 
   - If `-c` equals `True`(as a default value), that all rows will be colored despite the logic its given.
2. Using function `login()` it connects to the remote API to take `access_token` and then function `request_resourse(acc_tkn)` takes the resources located at `https://api.baubuddy.de/dev/index.php/v1/vehicles/select/active`.
3. Function `filter()` checks the data that was taken previously and searchs only for those resources that have a value set for `hu` field.
4. If `labelIds` was given as an input parameter to `-k`:
   - Function `get_color_code()` will connects to the API using  `https://api.baubuddy.de/dev/index.php/v1/labels/{labelId}` to take the value of the `ColorCode` for each record in our file and write it to the `labelIds` field.
   - However, if `labelIds` was not given as an input parameter, it won't show column `labelIds`
5. Using function `sort_excel()` all the data will be converted to xlsx file and will be sorted by field `gruppe`
6. If `-c` equals `True`, than function `color_rows()` will be active and compare the value in the filed `hu` to the folowing logic:
   - If `hu` is not older than 3 months --> green (`#007500`)
   - If `hu` is not older than 12 months --> orange (`#FFA500`)
   - If `hu` is older than 12 months --> red (`#b30000`)

As an output, [vehicles_2022-09-08_iso_formatted.xlsx](vehicles_2022-09-08_iso_formatted.xlsx) is given.
