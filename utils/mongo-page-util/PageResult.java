package com.th.trust.data.util;

import java.util.List;

/**
 * 分页结果.
 * 
 */
public class PageResult<T> {

//    @ApiModelProperty("页码，从1开始")
    private Integer pageNum;

//    @ApiModelProperty("页面大小")
    private Integer pageSize;

//    @ApiModelProperty("总数")
    private Long total;

//    @ApiModelProperty("总页数")
    private Integer pages;

//    @ApiModelProperty("数据")
    private List<T> list;

    public Integer getPageNum() {
        return pageNum;
    }

    public void setPageNum(Integer pageNum) {
        this.pageNum = pageNum;
    }

    public Integer getPageSize() {
        return pageSize;
    }

    public void setPageSize(Integer pageSize) {
        this.pageSize = pageSize;
    }

    public Long getTotal() {
        return total;
    }

    public void setTotal(Long total) {
        this.total = total;
    }

    public Integer getPages() {
        return pages;
    }

    public void setPages(Integer pages) {
        this.pages = pages;
    }

    public List<T> getList() {
        return list;
    }

    public void setList(List<T> list) {
        this.list = list;
    }

    @Override
    public String toString() {
        return "PageResult [pageNum=" + pageNum + ", pageSize=" + pageSize + ", total=" + total + ", pages=" + pages
                + ", list=" + list + "]";
    }

}