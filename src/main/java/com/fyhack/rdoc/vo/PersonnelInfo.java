package com.fyhack.rdoc.vo;

/**
 * PersonnelInfo
 * 干部人事档案专项审核情况登记表
 * <p/>
 *
 * @author elc_simayi
 * @since 2015/11/25
 */
public class PersonnelInfo {
    public String name; //姓名
    public String work_units_and_positions; //工作单位及职务
    public String work_level;  //级别
    public String other_opinion; //其他问题
    public String audit_opinion;  //审核意见

    public String getName() {
        int i = name.indexOf(":");
        if(i!=-1)
            name = name.substring(i+1);
        return name;
    }

    public String getWork_units_and_positions() {
        return work_units_and_positions;
    }

    public String getWork_level() {
        return work_level;
    }

    public String getAudit_opinion() {
        if(audit_opinion==null || audit_opinion.length()==0) {
            audit_opinion = "-";
        }
        return audit_opinion;
    }

    public String getOther_opinion() {
        if(other_opinion==null || other_opinion.length()==0) {
            other_opinion = "-";
        }
        return other_opinion;
    }
}
