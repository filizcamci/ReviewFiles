package apiReview;

import java.util.List;

import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

public class AnotherApiReview {

	public static void main(String[] args) {
	Response res=	RestAssured.given().get("https://www.obama.org/wp-json/wp/v2/users");
	JsonPath jp=res.jsonPath();
	System.out.println(jp.prettyPrint());
	List<String> names=jp.getList("name");
	List<Integer> ids=jp.getList("id");
	System.out.println(names);
	System.out.println(ids);
	List<String> avatarURLs=jp.getList("avatar_urls.24");
	System.out.println(avatarURLs);
	List<String> selfList=jp.getList("_links.self.href");
	System.out.println(selfList);
	for(int i=0;i<selfList.size();i++) {
		String self=selfList.get(i);
		//System.out.println(self.toString());
		
	}
	}
}
