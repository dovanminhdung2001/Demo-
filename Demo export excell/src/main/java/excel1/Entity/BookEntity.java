package excel1.Entity;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class BookEntity {
    private Long id;
    private String name;
    private Double price;
}
