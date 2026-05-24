# T1Sync C# — Visual Studio Setup

NuGet install ClosedXML

#

After fetching the seed's payload, we need replace field value with cell value.the replace logic is: 

For  a specific AttributeCode like 'Asset_Type', go through the meta, we can find all fields are grouped to  'level_1' and 'level_2, - indicating that under payload "AssetAttributes", fields are grouped to 2 nodes; 

  'level_1' means 'SearchPath' = 'Tree', has no '\\' - '\\' is the level separator, 

  'level_2' means  "SearchPath" = "Tree\\Street Tree", follows the above 'level_1' and one '\\' , there is a meaningful name 'Street Tree' 

 For field  'Near Power Line' , meta shows it is under 'asset_type' level 2, according to above logic, we can identify the single node ( which has "AttributeCode": "ASSET_TYPE" and "SearchPath": "Tree\Street Tree" under ' AssetAttributes' )   

Then meta data says it is 'Userfield1', then under the node ( identified by 'asset_type' and 'searchpath') , replace value of 'AttributeItemUserfield1'with the cell value