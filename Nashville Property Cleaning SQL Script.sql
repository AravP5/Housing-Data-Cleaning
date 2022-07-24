SELECT *
FROM master.dbo.Sheet1$; --General overview to note what columns and rows need cleaning

SELECT SaleDate, CONVERT(Date, SaleDate)
FROM master.dbo.Sheet1$

ALTER TABLE Sheet1$
ADD SaleDateFormatted DATE;

UPDATE Sheet1$
SET SaleDateFormatted = CONVERT(Date, SaleDate); -- New column SaleDateFormatted is populated with the formatted dates from SaleDate


SELECT *
FROM master.dbo.Sheet1$
WHERE PropertyAddress IS NULL -- Many values are null, and need to be populated

SELECT mastersheet.UniqueID,mastersheet.ParcelID, mastersheet.PropertyAddress, sheet2.UniqueID,sheet2.ParcelID, sheet2.PropertyAddress
FROM master.dbo.Sheet1$ mastersheet
LEFT JOIN master.dbo.Sheet1$ sheet2
    ON mastersheet.ParcelID = sheet2.ParcelID
    AND mastersheet.UniqueID != sheet2.UniqueID -- All the ParcelIDs are used to self join
WHERE mastersheet.PropertyAddress IS NULL -- We also filter out the Property Address to only see the NULLS which we have joined with the same ParcelIDs

UPDATE mastersheet
SET PropertyAddress = ISNULL(mastersheet.PropertyAddress, sheet2.PropertyAddress)
FROM master.dbo.Sheet1$ mastersheet
LEFT JOIN master.dbo.Sheet1$ sheet2
    ON mastersheet.ParcelID = sheet2.ParcelID
    AND mastersheet.UniqueID != sheet2.UniqueID
WHERE mastersheet.PropertyAddress IS NULL --Populates the NULL PropertyAddress column with the duplicate PropertyAddress column which has values
;--Rerunning to select for NULL PropertyAddress values yields no rows


SELECT PropertyAddress FROM master.dbo.Sheet1$ -- Need to split up the local address from the city address

SELECT 
SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1) AS Address,
SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) +1, LEN(PropertyAddress)) AS Address
FROM master.dbo.Sheet1$ -- Used SUBSTRING() and CHARINDEX() functions to separate the address +1 and -1 based on the comma

ALTER TABLE Sheet1$
ADD OnlyPropertyAddress VARCHAR(250),
    OnlyPropertyCity VARCHAR(250) --Added two new blank columns which will be populated with the local address and city, respectively
;

UPDATE Sheet1$
SET OnlyPropertyAddress = SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1), -- Split the address here and the line below
    OnlyPropertyCity = SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) +1, LEN(PropertyAddress));
;

SELECT OwnerAddress
FROM master.dbo.Sheet1$

SELECT PARSENAME(Replace(OwnerAddress, ',', '.'), 3),
PARSENAME(Replace(OwnerAddress, ',', '.'), 2),
PARSENAME(Replace(OwnerAddress, ',', '.'), 1)
FROM master.dbo.Sheet1$ -- PARSENAME() function an easier and faster way of splittings strings into substrings

ALTER TABLE Sheet1$
ADD OwnerLocalAddress VARCHAR(50),
    OwnerCityAddress VARCHAR(50),
    OwnerStateAddress VARCHAR(5) -- Created new blank columns which will be populated with substrings from OwnerAddress
;

UPDATE master.dbo.Sheet1$
SET OwnerLocalAddress = PARSENAME(Replace(OwnerAddress, ',', '.'), 3),
    OwnerCityAddress = PARSENAME(Replace(OwnerAddress, ',', '.'), 2),
    OwnerStateAddress = PARSENAME(Replace(OwnerAddress, ',', '.'), 1) -- Inserted substrings into new columns
;

SELECT DISTINCT SoldAsVacant, COUNT(SoldAsVacant)
FROM master.dbo.Sheet1$ 
GROUP BY SoldAsVacant
ORDER BY COUNT(SoldAsVacant) DESC -- Returns either N, No, Y, Yes. Needs to be either Yes or No

UPDATE Sheet1$
SET SoldAsVacant = CASE WHEN SoldAsVacant = 'Y' THEN 'Yes'
        WHEN SoldAsVacant = 'N' THEN 'No'
        ELSE SoldAsVacant
        END -- CASE statement used to change 'Y' and 'N' to the 'Yes' and 'No'
        ;


WITH RownumCTE AS( -- CTE created to and row_num > 1 to identify duplicates
    SELECT *,
    ROW_NUMBER() OVER (
        PARTITION BY ParcelID,
                    PropertyAddress,
                    SalePrice,
                    SaleDate,
                    LegalReference
                    ORDER BY UniqueID
                    )row_num
FROM master.dbo.Sheet1$
)

SELECT *
--DELETE -- Used to Delete the duplicate rows
FROM RownumCTE
WHERE row_num > 1
-- ORDER BY PropertyAddress


SELECT *
FROM master.dbo.Sheet1$

ALTER TABLE master.dbo.Sheet1$
DROP COLUMN OwnerAddress, PropertyAddress, SaleDate -- Dropped columns which no longer have any use
