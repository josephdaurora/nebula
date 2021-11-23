package daurora.nebula.app;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeRequestUrl;
import com.google.api.client.googleapis.auth.oauth2.GoogleTokenResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import javax.script.ScriptException;
import java.io.IOException;
import java.net.URISyntaxException;
import java.security.GeneralSecurityException;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static daurora.nebula.app.createTemplate.*;

public class analyzeData {

    private static final JsonFactory JSON_FACTORY = new JacksonFactory();
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static Credential credential;
    private static final HttpTransport HTTP_TRANSPORT = new NetHttpTransport();
    private static final String clientID = "YOUR ID HERE";
    private static final String clientSecret = "YOUR SECRET HERE";
    private static final String CALLBACK_URI = "YOUR CALLBACK URL HERE";


    private static String stateToken;
    private static GoogleAuthorizationCodeFlow flow;

    public analyzeData() {
        flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT,
                JSON_FACTORY, clientID, clientSecret, SCOPES).setAccessType("offline").build();

        generateStateToken();
    }

    public String buildLoginUrl() {

        final GoogleAuthorizationCodeRequestUrl url = flow.newAuthorizationUrl();

        return url.setRedirectUri(CALLBACK_URI).setState(stateToken).build();
    }

    static void generateStateToken(){

        SecureRandom sr1 = new SecureRandom();

        stateToken = "google;"+sr1.nextInt();

    }


    public static void wrapper(String authCode, String spreadsheetID, int numStudents, int numQuestions) throws GeneralSecurityException, IOException, URISyntaxException, ScriptException {
        GoogleTokenResponse tokenResponse = flow.newTokenRequest(authCode).setRedirectUri(CALLBACK_URI).execute();
        credential = flow.createAndStoreCredential(tokenResponse, null);

        removePreviousAnalyses(spreadsheetID);
        int gradedAssignmentWorksheetID = duplicateSpreadsheet(spreadsheetID,getWorksheetID(credential, spreadsheetID,1), 2);
        changeSheetTitle(credential, spreadsheetID, gradedAssignmentWorksheetID, "Graded Assignment");
        addScoringColumn(spreadsheetID, gradedAssignmentWorksheetID);
        coloredHighlight(spreadsheetID, gradedAssignmentWorksheetID, numStudents, numQuestions);
        formatAsPercentage(spreadsheetID,gradedAssignmentWorksheetID,4, numStudents+5, 4 ,5);
        gradeAssignment(spreadsheetID,"Graded Assignment", "Graded Assignment", numStudents,numQuestions, 0 , 0);
        int gradesSummaryWorksheetID = newWorksheet(spreadsheetID, "Grades Summary", 3, false);
        formatGradesSummary(spreadsheetID, gradesSummaryWorksheetID,numStudents,numQuestions);
        int mostMissedWorksheetID = newWorksheet(spreadsheetID, "Most Missed Questions", 4, false);
        formatMostMissed(spreadsheetID, mostMissedWorksheetID, numQuestions);
        int gradesSummarybyScoreWorksheetID = duplicateSpreadsheet(spreadsheetID,gradesSummaryWorksheetID, 5);
        changeSheetTitle(credential, spreadsheetID, gradesSummarybyScoreWorksheetID, "Grades Summary by Score");
        sortRange(spreadsheetID, gradesSummarybyScoreWorksheetID, 2, numStudents + 3, 1, 3, 2, "DESCENDING");
        int choiceAnalysisWorksheetID = newWorksheet(spreadsheetID, "Answer Choice Analysis", 6, false);
        choiceAnalysis(spreadsheetID, choiceAnalysisWorksheetID, numStudents, numQuestions);
    }


    public static String numberstoCoordinates(int columnNumber, int rowNumber) {


        StringBuilder result = new StringBuilder();

        while (columnNumber > 0) {
            int index = (columnNumber - 1) % 26;
            result.append((char) (index + 'A'));
            columnNumber = (columnNumber - 1) / 26;
        }

        result.append(rowNumber);
        return result.toString();

    }

    public static int duplicateSpreadsheet(String spreadsheetID, int startingSheetID, int numSheets) throws IOException {

        CopySheetToAnotherSpreadsheetRequest requestBody = new CopySheetToAnotherSpreadsheetRequest();
        requestBody.setDestinationSpreadsheetId(spreadsheetID);

        Sheets service = builderWithCredentials(credential);

        Sheets.Spreadsheets.SheetsOperations.CopyTo request =
                service.spreadsheets().sheets().copyTo(spreadsheetID, startingSheetID, requestBody);

        request.execute();

        return getWorksheetID(credential, spreadsheetID, numSheets);


    }

    public static int getWorksheetID(Credential credential, String spreadsheetID, int sheetPosition) throws IOException {

        Sheets service = builderWithCredentials(credential);

        Spreadsheet response1 = service.spreadsheets().get(spreadsheetID).setIncludeGridData(false)
                .execute();

        return response1.getSheets().get(sheetPosition).getProperties().getSheetId();
    }

    public static void coloredHighlight(String spreadsheetID, int worksheetID, int numStudents, int numQuestions) throws IOException {

        Sheets service = builderWithCredentials(credential);

        String incorrectFormula;
        String correctFormula;

        List<Request> requests = new ArrayList<>();

        int answerkeyIterator = 4;

        for (int columnsIterator = 6; columnsIterator <= numQuestions + 5; columnsIterator++) {

            for (int studentIterator = 5; studentIterator <= numStudents + 4; studentIterator++) {

                incorrectFormula = "=NE(" + "B" + answerkeyIterator + "," + numberstoCoordinates(columnsIterator, studentIterator) + ")";
                correctFormula = "=EQ(" + "B" + answerkeyIterator + "," + numberstoCoordinates(columnsIterator, studentIterator) + ")";

                requests.add(new Request().setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                        .setRule(new ConditionalFormatRule()
                                .setRanges(Collections.singletonList(new GridRange()
                                        .setSheetId(worksheetID)
                                        .setStartRowIndex(studentIterator - 1)
                                        .setEndRowIndex(studentIterator)
                                        .setStartColumnIndex(columnsIterator - 1)
                                        .setEndColumnIndex(columnsIterator)))

                                .setBooleanRule(new BooleanRule()
                                        .setCondition(new BooleanCondition()
                                                .setType("CUSTOM_FORMULA")
                                                .setValues(Collections.singletonList(
                                                        new ConditionValue().setUserEnteredValue(
                                                                incorrectFormula)
                                                ))
                                        )
                                        .setFormat(new CellFormat().setBackgroundColor(
                                                new Color().setRed(.918f).setGreen(0.6f).setBlue(0.6f)
                                        ))
                                )
                        )
                        .setIndex(0)));
                requests.add(new Request().setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                        .setRule(new ConditionalFormatRule()
                                .setRanges(Collections.singletonList(new GridRange()
                                        .setSheetId(worksheetID)
                                        .setStartRowIndex(studentIterator - 1)
                                        .setEndRowIndex(studentIterator)
                                        .setStartColumnIndex(columnsIterator - 1)
                                        .setEndColumnIndex(columnsIterator)))

                                .setBooleanRule(new BooleanRule()
                                        .setCondition(new BooleanCondition()
                                                .setType("CUSTOM_FORMULA")
                                                .setValues(Collections.singletonList(
                                                        new ConditionValue().setUserEnteredValue(
                                                                correctFormula)
                                                ))
                                        )
                                        .setFormat(new CellFormat().setBackgroundColor(
                                                new Color().setRed(.6f).setGreen(0.918f).setBlue(0.6f)
                                        ))
                                )
                        )
                        .setIndex(0)));
            }
            answerkeyIterator++;
        }
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();


    }

    public static void addScoringColumn(String spreadsheetID, int worksheetID) throws  IOException {

        Sheets service = builderWithCredentials(credential);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
                .setInsertDimension(new InsertDimensionRequest().setRange(new DimensionRange().setSheetId(worksheetID).setDimension("COLUMNS").setStartIndex(4).setEndIndex(5))));


        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(new GridRange()
                                .setSheetId(worksheetID)
                                .setStartColumnIndex(4)
                                .setEndColumnIndex(5)
                                .setStartRowIndex(3)
                                .setEndRowIndex(4))

                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue().setStringValue("Scores"))
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(12)
                                                .setBold(true)

                                        )
                                        .setHorizontalAlignment("CENTER")
                                )
                        )
                        .setFields("*")));


        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();

    }

    public static void gradeAssignment(String spreadsheetID, String sourceWorksheetName, String destinationWorksheetName, int numStudents, int numQuestions, int destinationColumnOffset, int destinationRowOffset) throws IOException {

        Sheets service = builderWithCredentials(credential);

        List<List<Object>> values;
        List<ValueRange> requests = new ArrayList<>();

        for (int i = 1; i <= numStudents; i++) {
            String formula = "=ROUND(100*(";

            for (int j = 6; j <= numQuestions + 5; j++) {
                formula += "EQ('" + sourceWorksheetName + "'!" + numberstoCoordinates(j, i + 4) + ",'" + sourceWorksheetName + "'!" + numberstoCoordinates(2, j - 2) + ")+";
            }

            if ((formula != null) && (formula.length() > 0)) {
                formula = formula.substring(0, formula.length() - 1);
                formula += ")/" + numQuestions + ")%";
            }

            values = Arrays.asList(Arrays.asList(formula));

            requests.add(new ValueRange()
                    .setRange("'" + destinationWorksheetName + "'!" + numberstoCoordinates(5 + destinationColumnOffset, i + 4 + destinationRowOffset))
                    .setValues(values));
        }
        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(requests);

        service.spreadsheets().values().batchUpdate(spreadsheetID, body).execute();

    }

    public static void formatAsPercentage(String spreadsheetID, int worksheetID, int startRow, int endRow, int startColumn, int endColumn) throws IOException {

        Sheets service = builderWithCredentials(credential);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest().setCell(new CellData().setUserEnteredFormat(new CellFormat().setNumberFormat(new NumberFormat().setType("PERCENT").setPattern("#,0%")))).setRange(new GridRange().setSheetId(worksheetID).setStartRowIndex(startRow).setEndRowIndex(endRow).setStartColumnIndex(startColumn).setEndColumnIndex(endColumn))
                        .setFields("*")));


        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();

    }

    public static int newWorksheet(String spreadsheetID, String sheetName, int numSheets, boolean isHidden) throws IOException{

        Sheets service = builderWithCredentials(credential);
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setAddSheet(new AddSheetRequest().setProperties(new SheetProperties().setTitle(sheetName).setHidden(isHidden))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);

        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();

        return getWorksheetID(credential, spreadsheetID, numSheets);
    }

    public static void formatGradesSummary(String spreadsheetID, int worksheetID, int numStudents, int numQuestions) throws  IOException {
        copyandPaste(credential, spreadsheetID, getWorksheetID(credential, spreadsheetID, 2), worksheetID, 4, numStudents + 4, 3, 4, 2, numStudents + 3, 1, 2, "PASTE_VALUES", "NORMAL");
        formatAsPercentage(spreadsheetID, worksheetID, 2, numStudents + 2, 2, 3);
        gradeAssignment(spreadsheetID, "Graded Assignment", "Grades Summary", numStudents, numQuestions, -2, -2);
        headingCells(credential, spreadsheetID, worksheetID, 0, 4, 0, 1, "MERGE_ROWS", 16, "Grades Summary");

    }

    public static void copyandPaste(Credential credential, String spreadsheetID, int originWorksheetID, int destinationWorksheetID, int or_startRow, int or_endRow, int or_startColumn, int or_endColumn, int dest_startRow, int dest_endRow, int dest_startColumn, int dest_endColumn, String pasteType, String pasteOrientation) throws IOException {

        Sheets service = builderWithCredentials(credential);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setCopyPaste(new CopyPasteRequest()
                .setSource(new GridRange().setSheetId(originWorksheetID).setStartRowIndex(or_startRow).setEndRowIndex(or_endRow).setStartColumnIndex(or_startColumn).setEndColumnIndex(or_endColumn))
                .setDestination(new GridRange().setSheetId(destinationWorksheetID).setStartRowIndex(dest_startRow).setEndRowIndex(dest_endRow).setStartColumnIndex(dest_startColumn).setEndColumnIndex(dest_endColumn)).setPasteType(pasteType).setPasteOrientation(pasteOrientation)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();
    }

    public static void formatMostMissed(String spreadsheetID, int worksheetID, int numQuestions) throws IOException {
        headingCells(credential, spreadsheetID, worksheetID, 0, 4, 0, 1, "MERGE_ROWS", 16, "Most Missed Questions");
        copyandPaste(credential, spreadsheetID, getWorksheetID(credential, spreadsheetID, 2), worksheetID, 3, numQuestions + 3, 0, 1, 2, numQuestions + 4, 1, 2, "PASTE_VALUES", "NORMAL");
        missedFrequency(spreadsheetID, "Graded Assignment", "Most Missed Questions", numQuestions);
        sortRange(spreadsheetID, worksheetID, 2, numQuestions + 3, 1, 3, 2, "DESCENDING");

    }

    public static void missedFrequency(String spreadsheetID, String sourceWorksheetName, String destinationWorksheetName, int numQuestions) throws IOException {

        Sheets service = builderWithCredentials(credential);

        List<List<Object>> values;
        List<ValueRange> requests = new ArrayList<>();

        for (int i = 6; i <= numQuestions + 5; i++) {
            String formula = "=" + numQuestions;

            formula += "-COUNTIF('" + sourceWorksheetName + "'!" + numberstoCoordinates(i, 5) + ":" + numberstoCoordinates(i, numQuestions + 4) + ",'" + sourceWorksheetName + "'!" + numberstoCoordinates(2, i - 2) + ")";

            formula = formula.substring(0, formula.length() - 1);
            formula += ")";


            values = Arrays.asList(Arrays.asList(formula));

            requests.add(new ValueRange()
                    .setRange("'" + destinationWorksheetName + "'!" + numberstoCoordinates(3, i - 3))
                    .setValues(values));
        }
        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(requests);
        service.spreadsheets().values().batchUpdate(spreadsheetID, body).execute();
    }

    public static void sortRange(String spreadsheetID, int worksheetID, int startRow, int endRow, int startColumn, int endColumn, int firstDimension, String sortOrder1) throws IOException {

        Sheets service = builderWithCredentials(credential);

        SortSpec sortSpec = new SortSpec();
        sortSpec.setSortOrder(sortOrder1).setDimensionIndex(firstDimension);

        List<Request> requests = new ArrayList<>();

        requests.add(new Request().setSortRange(new SortRangeRequest().setRange(new GridRange().setSheetId(worksheetID).setStartRowIndex(startRow).setEndRowIndex(endRow).setStartColumnIndex(startColumn).setEndColumnIndex(endColumn)).setSortSpecs(Collections.singletonList(sortSpec))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, body)
                .execute();
    }

    public static void choiceAnalysis(String spreadsheetID, int worksheetID, int numStudents, int numQuestions) throws IOException {
        int numAnswersPerQuestionWorksheetID = newWorksheet(spreadsheetID, "Number of Answers Per Question", 7, true);
        writeNumAnswersPerQuestion(spreadsheetID, numStudents, numQuestions);
        copyandPaste(credential, spreadsheetID, getWorksheetID(credential, spreadsheetID, 2), numAnswersPerQuestionWorksheetID, 3, numQuestions + 3, 0, 1, 2, numQuestions + 2, 1, 2, "PASTE_VALUES", "NORMAL");
        headingCells(credential, spreadsheetID,worksheetID,0,9,0,1,"MERGE_ROWS", 24,"Answer Choice Analysis");
        headingCells(credential, spreadsheetID, numAnswersPerQuestionWorksheetID, 0,4,0,1, "MERGE_ROWS", 16, "Number of Answers Per Question");


        Sheets service = builderWithCredentials(credential);


        formatAsPercentage(spreadsheetID, worksheetID, 0,1000,2,3);
        formatAsPercentage(spreadsheetID, worksheetID, 0,1000,6,7);

        int oddRow = 2;
        int evenRow = 2;
        List<List<Object>> answerChoicesValues, percentagesValues, sparklinesValues;
        List<ValueRange> requests = new ArrayList<>();
        int[] answersPerQuestion = readNumAnswersPerQuestion(spreadsheetID, numQuestions);

        boldCells(spreadsheetID, worksheetID, 1, 1000, 1, 2);
        boldCells(spreadsheetID, worksheetID, 1, 1000, 5, 6);

        List<Request> headingsRequest = new ArrayList<>();
        List<Request>  bordersRequest = new ArrayList<>();

        for(int i = 3; i < numQuestions+3; i++) {
            String formula = "=SORT(UNIQUE(UPPER('Graded Assignment'!" + numberstoCoordinates(i+3, 5)+ ":" + numberstoCoordinates(i+3, numQuestions+4) + "), FALSE, FALSE))";
            answerChoicesValues = Arrays.asList(Arrays.asList(formula));
            Border border = new Border().setColor(new Color().setRed(0f).setGreen(0f).setBlue(0f)).setStyle("SOLID_MEDIUM");

            if (i % 2 == 1) {
                requests.add(new ValueRange()
                        .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(2, oddRow + 2))
                        .setValues(answerChoicesValues));

                bordersRequest.add(new Request().setUpdateBorders(new UpdateBordersRequest().setRange(new GridRange()
                        .setSheetId(worksheetID)
                        .setStartRowIndex(oddRow)
                        .setEndRowIndex(oddRow+answersPerQuestion[i-3]+1)
                        .setStartColumnIndex(1)
                        .setEndColumnIndex(4))
                        .setTop(border)
                        .setBottom(border)
                        .setLeft(border)
                        .setRight(border)));


                for (int j = 0; j< answersPerQuestion[i-3] ; j++) {
                    String percentages = "=COUNTIF('Graded Assignment'!" + numberstoCoordinates(i+3, 5) + ":" + numberstoCoordinates(i+3, numQuestions+4) + "," + numberstoCoordinates(2, oddRow+2+j) + ")/" + numStudents;
                    String sparklines = "=SPARKLINE(" + numberstoCoordinates(2,oddRow + 2 + j) + ":" + numberstoCoordinates(3,oddRow + 2 + j) + ",{" + '"'+ "charttype" + '"' +  "," + '"' + "bar" + '"' + ";" + '"' + "max" + '"' + "," + "1})";
                    percentagesValues = Arrays.asList(Arrays.asList(percentages));
                    sparklinesValues = Arrays.asList(Arrays.asList(sparklines));
                    requests.add (new ValueRange()
                            .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(3, oddRow + 2 + j))
                            .setValues(percentagesValues));
                    requests.add (new ValueRange()
                            .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(4, oddRow + 2 + j))
                            .setValues(sparklinesValues));

                    headingsRequest.add( new Request()
                            .setMergeCells(new MergeCellsRequest()
                                    .setRange(new GridRange()
                                            .setSheetId(worksheetID)
                                            .setStartColumnIndex(1)
                                            .setEndColumnIndex(4)
                                            .setStartRowIndex(oddRow)
                                            .setEndRowIndex(oddRow + 1))
                                    .setMergeType("MERGE_ROWS")

                            ));

                    headingsRequest.add(new Request()
                            .setRepeatCell(new RepeatCellRequest()
                                    .setRange(new GridRange()
                                            .setSheetId(worksheetID)
                                            .setStartColumnIndex(1)
                                            .setEndColumnIndex(4)
                                            .setStartRowIndex(oddRow)
                                            .setEndRowIndex(oddRow + 1))
                                    .setCell(new CellData()
                                            .setUserEnteredValue(new ExtendedValue().setStringValue("Question " + (i-2)))
                                            .setUserEnteredFormat(new CellFormat()
                                                    .setTextFormat(new TextFormat()
                                                            .setFontSize(12)
                                                            .setBold(Boolean.TRUE)

                                                    )
                                                    .setHorizontalAlignment("CENTER")
                                            )
                                    )

                                    .setFields("*")

                            )
                    );

                }
                oddRow+= answersPerQuestion[i-3] + 2 ;

            }

            if (i % 2 == 0) {
                requests.add(new ValueRange()
                        .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(6, evenRow + 2))
                        .setValues(answerChoicesValues));

                bordersRequest.add(new Request().setUpdateBorders(new UpdateBordersRequest().setRange(new GridRange()
                        .setSheetId(worksheetID)
                        .setStartRowIndex(evenRow)
                        .setEndRowIndex(evenRow+answersPerQuestion[i-3]+1)
                        .setStartColumnIndex(5)
                        .setEndColumnIndex(8))
                        .setTop(border)
                        .setBottom(border)
                        .setLeft(border)
                        .setRight(border)));

                for (int j = 0; j<=3 ; j++) {
                    String percentages = "=COUNTIF('Graded Assignment'!" + numberstoCoordinates(i+3, 5) + ":" + numberstoCoordinates(i+3, numQuestions+4) + "," + numberstoCoordinates(2, evenRow+2+j) + ")/" + numStudents;
                    String sparklines = "=SPARKLINE(" + numberstoCoordinates(6,evenRow + 2 + j) + ":" + numberstoCoordinates(7,evenRow + 2 + j) + ",{" + '"'+ "charttype" + '"' +  "," + '"' + "bar" + '"' + ";" + '"' + "max" + '"' + "," + "1})";
                    percentagesValues = Arrays.asList(Arrays.asList(percentages));
                    sparklinesValues = Arrays.asList(Arrays.asList(sparklines));
                    requests.add (new ValueRange()
                            .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(7, evenRow + 2 + j))
                            .setValues(percentagesValues));
                    requests.add (new ValueRange()
                            .setRange("'Answer Choice Analysis'!" + numberstoCoordinates(8, evenRow + 2 + j))
                            .setValues(sparklinesValues));

                    headingsRequest.add( new Request()
                            .setMergeCells(new MergeCellsRequest()
                                    .setRange(new GridRange()
                                            .setSheetId(worksheetID)
                                            .setStartColumnIndex(5)
                                            .setEndColumnIndex(8)
                                            .setStartRowIndex(evenRow)
                                            .setEndRowIndex(evenRow + 1))
                                    .setMergeType("MERGE_ROWS")

                            ));

                    headingsRequest.add(new Request()
                            .setRepeatCell(new RepeatCellRequest()
                                    .setRange(new GridRange()
                                            .setSheetId(worksheetID)
                                            .setStartColumnIndex(5)
                                            .setEndColumnIndex(8)
                                            .setStartRowIndex(evenRow)
                                            .setEndRowIndex(evenRow + 1))
                                    .setCell(new CellData()
                                            .setUserEnteredValue(new ExtendedValue().setStringValue("Question " + (i-2)))
                                            .setUserEnteredFormat(new CellFormat()
                                                    .setTextFormat(new TextFormat()
                                                            .setFontSize(12)
                                                            .setBold(Boolean.TRUE)

                                                    )
                                                    .setHorizontalAlignment("CENTER")
                                            )
                                    )

                                    .setFields("*")

                            )
                    );

                }

                evenRow+= answersPerQuestion[i-3] + 2 ;

                if (oddRow > evenRow)
                {
                    evenRow = oddRow;
                }

                else if(evenRow > oddRow)
                {
                    oddRow = evenRow;
                }

            }



        }

        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
                    .setValueInputOption("USER_ENTERED")
                    .setData(requests);
        service.spreadsheets().values().batchUpdate(spreadsheetID, body).execute();

        BatchUpdateSpreadsheetRequest headingsBody = new BatchUpdateSpreadsheetRequest()
                .setRequests(headingsRequest);
        service.spreadsheets()
                .batchUpdate(spreadsheetID, headingsBody)
                .execute();


        BatchUpdateSpreadsheetRequest bordersBody = new BatchUpdateSpreadsheetRequest().setRequests(bordersRequest);
        service.spreadsheets().batchUpdate(spreadsheetID, bordersBody).execute();



    }

    public static void writeNumAnswersPerQuestion(String spreadsheetID, int numStudents, int numQuestions) throws IOException {

        Sheets service = builderWithCredentials(credential);

        List<List<Object>> values;
        List<ValueRange> requests = new ArrayList<>();

        for (int i = 6; i <= numQuestions + 5; i ++)
        {
            String formula = "=COUNTUNIQUE(UPPER('Graded Assignment'!" + numberstoCoordinates(i,5) + ")";
            for (int j = 6; j <= numStudents + 4; j++){
                formula += ",UPPER('Graded Assignment'!" + numberstoCoordinates(i, j) + ")";
            }
            formula+= ")";
            values = Arrays.asList(Arrays.asList(formula));
            requests.add(new ValueRange()
                    .setRange("'Number of Answers Per Question'!" + numberstoCoordinates(3, i-3))
                    .setValues(values));

        }
        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(requests);
        service.spreadsheets().values().batchUpdate(spreadsheetID, body).execute();
    }

    private static int[] readNumAnswersPerQuestion(String spreadsheetID, int numQuestions) throws IOException {

        Sheets service = builderWithCredentials(credential);

        ValueRange result = service.spreadsheets().values().get(spreadsheetID,"'Number of Answers Per Question'!C3:C" + (numQuestions + 2)).execute();
        Object[] arrayResult = result.getValues().toArray();

        int[] finalResult = new int[arrayResult.length];

        for(int i = 0; i<= arrayResult.length-1; i++){

            finalResult[i] = Integer.parseInt(arrayResult[i].toString().substring(1,arrayResult[i].toString().length()-1));
            System.out.println(arrayResult.length + " | " + finalResult[i]);
        }

        return finalResult;

    }

    public static void boldCells (String spreadsheetID, int worksheetID, int startRow, int endRow, int startColumn, int endColumn) throws IOException {

        Sheets service = builderWithCredentials(credential);

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
                                .setUserEnteredFormat(new CellFormat()
                                        .setTextFormat(new TextFormat()
                                                .setFontSize(10)
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

    public static void removePreviousAnalyses(String spreadsheetID) throws IOException {
        Sheets service = builderWithCredentials(credential);


        Spreadsheet request = service.spreadsheets().get(spreadsheetID).execute();
        List<Sheet> listofSheets = request.getSheets();
        List<String> sheetNames = Arrays.asList("Graded Assignment", "Grades Summary", "Most Missed Questions", "Grades Summary by Score", "Answer Choice Analysis", "Number of Answers Per Question");


        for(int i = 0; i < listofSheets.size(); i++)
        {
            for (int j = 0; j < sheetNames.size(); j++)
            {

                if (listofSheets.get(i).getProperties().getTitle().equalsIgnoreCase(sheetNames.get(j))) {
                    deleteSheet(spreadsheetID, listofSheets.get(i).getProperties().getSheetId());
                }
            }

        }

    }

    public static void deleteSheet(String spreadsheetID, int worksheetID) throws IOException {
        Sheets service = builderWithCredentials(credential);


        List<Request> requests = new ArrayList<>();

        requests.add(new Request().setDeleteSheet(new DeleteSheetRequest().setSheetId(worksheetID)));
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetID, body).execute();
    }


}

