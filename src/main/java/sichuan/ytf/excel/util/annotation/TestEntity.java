package sichuan.ytf.excel.util.annotation;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class TestEntity {
    @Excel(name = "乙方联系电话")
    private String secondContactNumber;

    // 表单状况 ConfirmationRightStateType
    @Excel(name = "撒地方", readConverterExp = "0=原始,1=变更")
    private Integer crState;

}
