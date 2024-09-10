package Pages;

import java.util.HashMap;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;




public class DemoAPI {
	


	public String setPlaceHolder(String sheetName) throws JsonProcessingException {
		
		HashMap<String,Object>Mk= new HashMap<String,Object>();
		HashMap<String,Object>sk= new HashMap<String,Object>();
		Mk.put("name", "Apple MacBook Pro 16");
		sk.put("year", "2019");
		sk.put("price", "1849.99");
		sk.put("CPU model","Intel Core i9");//input
		sk.put("Hard disk size","1 TB");
		Mk.put("data", sk);

		String json = new ObjectMapper().writeValueAsString(Mk);
		return json;
	}

}
