package excel1.pmud.model.req;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class ChangePasswordReq {
    private String oldPass;
    private String newPass;
    private String repeat;
}
