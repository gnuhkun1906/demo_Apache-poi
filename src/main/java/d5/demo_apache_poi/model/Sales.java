package d5.demo_apache_poi.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;

@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Sales {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    private String nganhHang;
    private int quanAo;
    private int giayDep;
    private int tuiSach;
    private int muNon;

    public Sales(String nganhHang, int quanAo, int giayDep, int tuiSach, int muNon) {
        this.nganhHang=nganhHang;
        this.giayDep=giayDep;
        this.quanAo=quanAo;
        this.tuiSach=tuiSach;
        this.muNon=muNon;
    }
}
