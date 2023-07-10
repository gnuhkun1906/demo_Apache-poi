package d5.demo_apache_poi.service;

import d5.demo_apache_poi.model.Sales;
import d5.demo_apache_poi.repository.ISales;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
@RequiredArgsConstructor
public class SaleService {
    private final ISales saleRepository;
    public void save (Sales sales) {
        saleRepository.save(sales);
    }
    public List<Sales> findAll() {
      return saleRepository.findAll();
    }

}
