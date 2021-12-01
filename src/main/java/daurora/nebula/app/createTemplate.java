package daurora.nebula.app;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeRequestUrl;
import com.google.api.client.googleapis.auth.oauth2.GoogleTokenResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import java.io.IOException;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static daurora.nebula.app.analyzeData.getWorksheetID;


public class createTemplate {

    private static final JsonFactory JSON_FACTORY = new JacksonFactory();
    private static final HttpTransport HTTP_TRANSPORT = new NetHttpTransport();
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);

    private static final String clientID = "YOUR ID HERE";
    private static final String clientSecret = "YOUR SECRET HERE";
    private static final String CALLBACK_URI = "YOUR CALLBACK URL HERE";

    private static Credential credential;

    private final GoogleAuthorizationCodeFlow flow;

    public createTemplate() {
        flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT,
                JSON_FACTORY, clientID, clientSecret, SCOPES).setAccessType("offline").build();

    }

    public String buildLoginUrl() {
        final GoogleAuthorizationCodeRequestUrl url = flow.newAuthorizationUrl();
        SecureRandom sr1 = new SecureRandom();
        String stateToken = "google;" + sr1.nextInt();
        return url.setRedirectUri(CALLBACK_URI).setState(stateToken).build();
    }


    public static Sheets builder() {
        return new Sheets.Builder(HTTP_TRANSPORT, GsonFactory.getDefaultInstance(), credential )
                .setApplicationName("Nebula")
                .build();
    }

    public static Sheets builderWithCredentials(Credential providedCredential) {
        return new Sheets.Builder(HTTP_TRANSPORT, GsonFactory.getDefaultInstance(), providedCredential )
                .setApplicationName("Nebula")
                .build();
    }

    public String createSpreadsheet(String authCode, String spreadsheetTitle, int numStudents, int numQuestions) throws IOException {

        GoogleTokenResponse tokenResponse = flow.newTokenRequest(authCode).setRedirectUri(CALLBACK_URI).execute();
        credential = flow.createAndStoreCredential(tokenResponse, null);

        Sheets service = builder();

        Spreadsheet spreadsheet = new Spreadsheet()
                .setProperties(new SpreadsheetProperties()
                        .setTitle(spreadsheetTitle));

        spreadsheet = service.spreadsheets().create(spreadsheet)
                .setFields("spreadsheetId")
                .execute();

        String spreadsheetID = spreadsheet.getSpreadsheetId();
        dashboard(credential, spreadsheetID, 0, numStudents, numQuestions);
        int studentDataSheetID = newWorksheet(spreadsheetID, "Student Data", 1);
        formatStudentData(spreadsheetID, studentDataSheetID, spreadsheetTitle, numStudents, numQuestions);
        return spreadsheetID;
    }

    public void dashboard(Credential credential, String spreadsheetID, int worksheetID, int numStudents, int numQuestions) throws IOException {

        changeSheetTitle(credential, spreadsheetID,0, "Dashboard");
        headingCells(credential, spreadsheetID, worksheetID, 0,9,0,1, "MERGE_ROWS", 24, "Welcome to your Nebula Template!");
        mergeCells(credential, spreadsheetID,worksheetID,0,9,2,3, "MERGE_ROWS");
        mergeCells(credential, spreadsheetID, worksheetID, 0,9,4,5, "MERGE_ROWS");
        mergeCells(credential, spreadsheetID, worksheetID, 0,9,6, 7, "MERGE_ROWS");


        Sheets service = builder();

        String dashboardText1 = "To begin, go to the Student Data sheet and fill in the answer key and your student grades";
        String dashboardText2 = "Then, come back here and click the Analyze button below. You may be asked to authenticate with Google to ensure access to this document.";
        String dashboardText3 = "If you ever need to reanalyze the assignment, click the Analyze button again and all sheets except Student Data will be deleted and regenerated with the updated information.";


        List<Request> requests = new ArrayList<>();

        bulkAddText(requests, worksheetID, 0,9,2,3,12, dashboardText1, "WRAP");
        bulkAddText(requests, worksheetID, 0,9,4,5,12, dashboardText2, "WRAP");
        bulkAddText(requests, worksheetID, 0,9,6,7,12, dashboardText3, "WRAP");

        String formula = "=HYPERLINK(" + '"' + "  https://YOUR URL HERE/analyze?spreadsheetID=" + spreadsheetID + "&numStudents=" + numStudents + "&numQuestions=" + numQuestions + '"' + "," + '"' + "Analyze" + '"' + ")";


        newButton(requests, spreadsheetID, worksheetID,4,6,9,11,18, formula, .6f,0f,1f,1f,1f,1f);

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();

    }

    public static void changeSheetTitle(Credential credential, String spreadsheetID, int worksheetID, String newTitle) throws IOException {

        Sheets service = builderWithCredentials(credential);
        List<Request> requests = new ArrayList<>();

        requests.add( new Request()
                .setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
                        .setProperties( new SheetProperties()
                                .setSheetId(worksheetID)
                                .setTitle(newTitle))
                        .setFields("title")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);

        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();


    }

    public static void headingCells(Credential credential, String spreadsheetID, int worksheetID, int startColumn, int endColumn, int startRow, int endRow, String mergeType, int textSize, String data) throws IOException{


        Sheets service = builderWithCredentials(credential);


        mergeCells(credential, spreadsheetID,worksheetID, startColumn, endColumn, startRow, endRow, mergeType);

        List<Request> requests = new ArrayList<>();


        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(startColumn)
                                .setEndColumnIndex(endColumn)
                                .setStartRowIndex(startRow)
                                .setEndRowIndex(endRow))
                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue().setStringValue(data))
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(textSize)
                                                .setBold(Boolean.TRUE)

                                        )
                                        .setHorizontalAlignment("CENTER")
                                )
                        )

                        .setFields("*")

                )
        );


        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();

    }

    public static void mergeCells(Credential credential, String spreadsheetID, int worksheetID, int startColumn, int endColumn, int startRow, int endRow, String mergeType) throws IOException {


        Sheets service = builderWithCredentials(credential);

        List<Request> requests = new ArrayList<>();

        requests.add( new Request()
                .setMergeCells(new MergeCellsRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(startColumn)
                                .setEndColumnIndex(endColumn)
                                .setStartRowIndex(startRow)
                                .setEndRowIndex(endRow))
                        .setMergeType(mergeType)

                ));


        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();


    }

    public static void bulkAddText (List<Request> requests, int worksheetID, int startColumn, int endColumn, int startRow, int endRow, int textSize, String data, String wrapStrategy)  {

        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(startColumn)
                                .setEndColumnIndex(endColumn)
                                .setStartRowIndex(startRow)
                                .setEndRowIndex(endRow))
                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue().setStringValue(data))
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(textSize)

                                        )
                                        .setHorizontalAlignment("CENTER")
                                        .setWrapStrategy(wrapStrategy)
                                )
                        )

                        .setFields("*")

                )
        );

    }

    public void newButton(List <Request> requests, String spreadsheetID, int worksheetID, int startColumn, int endColumn, int startRow, int endRow, int textSize, String data, float backgroundRed, float backgroundGreen, float backgroundBlue, float textRed, float textGreen, float textBlue) throws  IOException {

        mergeCells(credential, spreadsheetID,worksheetID,startColumn, endColumn, startRow, endRow, "MERGE_ALL");


        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(startColumn)
                                .setEndColumnIndex(endColumn)
                                .setStartRowIndex(startRow)
                                .setEndRowIndex(endRow))
                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue().setFormulaValue(data))
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(textSize)
                                                .setBold(Boolean.TRUE)
                                                .setForegroundColor(new Color().setRed(textRed).setGreen(textGreen).setBlue(textBlue)))
                                        .setBackgroundColor(new Color().setRed(backgroundRed).setGreen(backgroundGreen).setBlue(backgroundBlue))
                                        .setHorizontalAlignment("CENTER")
                                        .setVerticalAlignment("MIDDLE")

                                )
                        )

                        .setFields("*")

                )
        );

    }

    public  int newWorksheet(String spreadsheetID, String title, int index) throws IOException {

        Sheets service = builder();


        List<Request> requests = new ArrayList<>();

        requests.add(new Request().setAddSheet(new AddSheetRequest().setProperties(new SheetProperties().setTitle(title).setIndex(index))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
         service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();
        return getWorksheetID(credential, spreadsheetID,index);

    }

    public static void formatStudentData(String spreadsheetID, int worksheetID, String spreadsheetTitle, int numStudents, int numQuestions) throws IOException {


        Sheets service = builder();


        headingCells(credential, spreadsheetID, worksheetID,0, 2, 2, 3, "MERGE_ROWS", 16, "Answer Key");
        appendColumns(spreadsheetID, numQuestions);
        headingCells(credential, spreadsheetID, worksheetID,3, numQuestions+4, 2, 3, "MERGE_ROWS",16, "Student Scores");
        headingCells(credential, spreadsheetID, worksheetID,0, numQuestions+4, 0, 1, "MERGE_ROWS",36, spreadsheetTitle);

        List<Request> requests = new ArrayList<>();


        for (int i = 4; i <= numQuestions +3 ; i++)
        {
            bulkCenteredCells(requests, worksheetID, 0, 1, i-1, i, 12, "Question ", 3, false);
            bulkCenteredCells(requests, worksheetID, i, i+1, 3, 4, 12, "Question ", 4, true);
        }

        for (int i = 4; i <= numStudents +3 ; i++)
        {

            bulkCenteredCells(requests, worksheetID,3, 4, i, i+1, 12, "Student ",4,  false);
        }

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);

        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();

    }

    public static void bulkCenteredCells(List<Request> queuedRequests, int worksheetID,  int startColumn, int endColumn, int startRow, int endRow, int textSize, String prefix, int offset, Boolean inColumns){


        int finalSuffix;
        if (inColumns) {

            finalSuffix = endColumn - offset;

        }

        else {
            finalSuffix = endRow - offset;
        }
        queuedRequests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(startColumn)
                                .setEndColumnIndex(endColumn)
                                .setStartRowIndex(startRow)
                                .setEndRowIndex(endRow))

                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue().setStringValue(prefix + finalSuffix))
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(textSize)

                                        )
                                        .setHorizontalAlignment("CENTER")
                                )
                        )
                        .setFields("*")));


    }

    public static void appendColumns(String spreadsheetID, int numColumns) throws IOException {

        Sheets service = builder();

        List<Request> requests = new ArrayList<>();

        requests.add(new Request ()
                .setAppendDimension(new AppendDimensionRequest()
                        .setDimension("COLUMNS")
                        .setLength(numColumns)
                )
        );

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();


    }

    public static String spreadsheetURL(String spreadsheetID) {
        return "https://docs.google.com/spreadsheets/d/" + spreadsheetID + "/edit#gid=0";
    }
}

