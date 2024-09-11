package TESTRUNNER;

import io.restassured.RestAssured;

import io.restassured.http.ContentType;
import io.restassured.response.Response;
public class APIRequest_Response {

	public static Response POSTRequest(String Requestbody,String url) {
		
		Response response=RestAssured.given().contentType(ContentType.JSON).accept(ContentType.JSON).
			    body(Requestbody).when().log().all().post(url);
				
		return response;
		
	}
public static Response GETRequest(String url) {
		
		Response response=RestAssured.given().contentType(ContentType.JSON).accept(ContentType.JSON).when().log().all().get(url);
				
		return response;
		
	}
}
