package com.bandar.excel.Entity;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;

@Controller
public class CustomerController {

	@Autowired private CustomerRepository repo;
	
	@GetMapping("/customer")
	public String getCustomer(Model theModel) {
		
		theModel.addAttribute("customers", repo.findAll());
		return "customer";
	}
	
	@PostMapping("/customer/addNew")
	public String create(Customer customer) {
		repo.save(customer);
		return "redirect:/customer";
	}
	
	@GetMapping("/users/export/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        											  //"yyyy-MM-dd_HH:mm:ss"
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd");
        String currentDateTime = dateFormatter.format(new Date());
         
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Customer_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);
         
        List<Customer> listCustomers = repo.findAll();
         
        CustomerExcelExporter excelExporter = new CustomerExcelExporter(listCustomers);
         
        excelExporter.export(response);    
    } 
}
