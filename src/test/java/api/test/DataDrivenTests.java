package api.test;

import api.endpoints.UserEndPoints;
import api.payload.User;
import api.utilities.DataProviders;
import io.restassured.response.Response;

import org.testng.Assert;
import org.testng.annotations.Test;

public class DataDrivenTests {

	@Test(priority=1, dataProvider="Data", dataProviderClass=DataProviders.class)
	public void testPostUser(String userID, String userName, String fname, String lname, String userEmail, String pass, String ph)
	{
		User userPayload = new User();
		userPayload.setId(Integer.parseInt(userID));
		userPayload.setUsername(userName);
		userPayload.setFirstName(fname);
		userPayload.setLastName(lname);
		userPayload.setEmail(userEmail);
		userPayload.setPassword(pass);
		userPayload.setPhone(ph);
		
		Response response = UserEndPoints.createUser(userPayload);
		Assert.assertEquals(response.getStatusCode(), 200);
		
	}
	
	@Test(priority=2, dataProvider="UserNames", dataProviderClass=DataProviders.class)
	public void testDeleteUserByName(String userName) {
		Response response = UserEndPoints.deleteUser(userName);
		Assert.assertEquals(response.getStatusCode(), 200);
	}
}
