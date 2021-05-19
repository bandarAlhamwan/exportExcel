package com.bandar.excel.Entity;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class CusotmerRestController {
	
	@Autowired private CustomerRepository repo;

	@GetMapping("/customerr")
	public List<Customer> getCusotmer(){
		return repo.findAll();
	}
}
