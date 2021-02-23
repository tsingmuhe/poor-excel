package com.sunchangpeng.poor.excel.sax;

@FunctionalInterface
public interface RowFilter {
    boolean test(Integer t, Integer u);
}
