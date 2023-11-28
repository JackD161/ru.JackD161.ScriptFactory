public class ClearMark {
    private String data;
    public ClearMark(String path, String code) {
        data = "Update items\nset HAS_TAGS=0\nwhere IID=" + code + ";\nexecute procedure PR_SCRIPT_STAGE(777, '{86D544A3-DF27-4922-9A7F-E5E3836A47D5}', 0, 'ItemsHasTagsClear_" + code + ".sql', '" + code + "', 0);\nCOMMIT;";
        new SaveFile(path, "ItemsHasTagsClear_" + code + ".sql", data);
    }
}
