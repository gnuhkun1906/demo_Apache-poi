package d5.demo_apache_poi.repository;

import d5.demo_apache_poi.model.Sales;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ISales extends JpaRepository<Sales,Long> {

}
