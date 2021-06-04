package de.thkoeln;

import org.apache.poi.ss.usermodel.Row;

public class SortRow
{
    public SortRow()
    {
    }

    public SortRow(String key, Row value)
    {
        this.Key = key;
        this.Value = value;
    }

    private String Key;
    private Row Value;

    public String getKey() {
        return Key;
    }

    public void setKey(String key) {
        Key = key;
    }

    public Row getValue() {
        return Value;
    }

    public void setValue(Row value) {
        Value = value;
    }
}
