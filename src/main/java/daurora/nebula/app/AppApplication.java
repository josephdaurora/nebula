package daurora.nebula.app;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.view.RedirectView;

import javax.script.ScriptException;
import java.io.IOException;
import java.net.URISyntaxException;
import java.security.GeneralSecurityException;

import static daurora.nebula.app.createTemplate.spreadsheetURL;

@SpringBootApplication
@Controller
public class AppApplication {
	private final createTemplate templateHelper = new createTemplate();
	private final analyzeData analysisHelper = new analyzeData();

	private String storedAssignmentName;
	private int storednumStudents;
	private int storednumQuestions;
	private String storedSpreadsheetID;

	public static void main(String[] args) {
		SpringApplicationBuilder builder = new SpringApplicationBuilder(AppApplication.class);
		builder.headless(false).run(args);
	}

	@GetMapping("/createtemplate")
	@ResponseBody
	public RedirectView createTemplate(@RequestParam(name="assignmentName") String assignmentName, @RequestParam(name="numStudents") int numStudents, @RequestParam(name="numQuestions") int numQuestions,Model model)  {
		model.addAttribute("assignmentName", assignmentName);
		model.addAttribute("numStudents", numStudents);
		model.addAttribute("numQuestions", numQuestions);

		storedAssignmentName = assignmentName;
		storednumStudents = numStudents;
		storednumQuestions = numQuestions;

		String googleURL = templateHelper.buildLoginUrl();

		return new RedirectView(googleURL);
	}

	@GetMapping("/templateoauth")
	public String templateOAuth(@RequestParam(name="code") String authCode, @RequestParam(name="state") String state, Model model) throws IOException, ScriptException, GeneralSecurityException, URISyntaxException {

		if (authCode != null) {
			String spreadsheetID = templateHelper.createSpreadsheet(authCode, storedAssignmentName, storednumStudents, storednumQuestions);
			model.addAttribute("spreadsheetURL", spreadsheetURL(spreadsheetID));
			congratulations(spreadsheetURL(spreadsheetID), model);
			return "../static/congratulations";
		}
		else
		{
			return "error";
		}
	}

	@GetMapping("/greeting")
	public String greeting(@RequestParam(name = "name", required = false, defaultValue = "World") String name, Model model) {
		model.addAttribute("name", name);
		return "greeting";
	}

	@GetMapping("/error")
	public String error() {
		return "error";
	}

	@GetMapping("/analysisoauth")
	public String analysisOAuth(@RequestParam(name="state") String state, @RequestParam(name="code") String authCode,  Model model) throws IOException, ScriptException, GeneralSecurityException, URISyntaxException {

		if (authCode != null) {
			analysisHelper.wrapper(authCode, storedSpreadsheetID, storednumStudents, storednumQuestions);
			analysiscomplete(spreadsheetURL(storedSpreadsheetID), model);
			return "../static/analysiscomplete";

		}
		else
		{
			return "error";
		}
	}

	@GetMapping("/analyze")
	public RedirectView analyze(@RequestParam(name = "spreadsheetID") String spreadsheetID,@RequestParam(name = "numStudents") int numStudents, @RequestParam(name = "numQuestions") int numQuestions, Model model) {

		model.addAttribute("spreadsheetID", spreadsheetID);
		model.addAttribute("numStudents", numStudents);
		model.addAttribute("numQuestions", numQuestions);

		storedSpreadsheetID = spreadsheetID;
		storednumStudents = numStudents;
		storednumQuestions = numQuestions;

		String googleURL = analysisHelper.buildLoginUrl();

		return new RedirectView(googleURL);

	}

	@GetMapping("../static/congratulations")
	public String congratulations(@RequestParam(name = "spreadsheetURL", required = true) String spreadsheetURL, Model model) {
		model.addAttribute("spreadsheetURL", spreadsheetURL);
		return "../static/congratulations";
	}

	@GetMapping("../static/analysiscomplete")
	public String analysiscomplete(@RequestParam(name = "spreadsheetURL", required = true) String spreadsheetURL, Model model) {
		model.addAttribute("spreadsheetURL", spreadsheetURL);
		return "../static/congratulations";
	}


}