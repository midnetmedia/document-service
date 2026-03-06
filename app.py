"""
Document Processing Service for Sure Forms
"""

from flask import Flask, request, send_file, jsonify
import zipfile
import io
import os
import base64
import re
from datetime import datetime

app = Flask(__name__)

# Field mapping from Sure Forms to Word placeholders
FIELD_MAPPING = {
    "1-what-is-the-name-of-your-school": "SchoolName",
    "2-what-is-the-name-of-your-school-district-or-management-organization": "SchoolDistrict",
    "3-it-is-a-public-charter-or-private-school": "SchoolType",
    "4-your-school-is-located-in-which-city": "SchoolLocation",
    "5-your-school-is-located-in-which-state": "SchoolState",
    "6-what-is-the-year-your-school-was-founded": "SchoolYear",
    "7-how-many-students-attend-your-school": "StudentNumber",
    "8-what-level-of-school-serves-the-students": "ServiceLevels",
    "9-what-grades-does-your-school-address": "SchoolGrades",
    "10-what-is-the-mission-of-your-school": "SchoolMission",
    "11-who-leads-your-school": "SchoolHead",
    "11a-the-title-of-the-person-who-leads-our-school-is": "SchoolHeadTitle",
    "12-is-your-school-principal-supportive-of-securing-grant-funds-for-school-and-classroom-projects": "FunderSupport",
    "13-what-are-the-strengths-or-uniqueness-of-your-school": "StrengthsWeaknesses",
    "14-describe-the-top-educational-projects-in-your-school-that-are-aligned-with-your-schools-mission": "TopProjects",
    "15-what-is-your-website-address": "SchoolURL",
    "16-describe-the-classroom-environment-you-foster-and-the-grades-you-teach": "TeacherGradeLevel",
    "17-what-are-some-interesting-or-unique-qualities-about-the-students-in-your-classroom": "StudentDescription",
    "18-what-is-the-fundamental-issue-or-challenge-that-you-aim-to-address-with-your-students": "TeacherType",
    "19-why-is-the-issue-or-challenge-happening-be-clear-about-it-estimate-percentages": "WhyTheProblem",
    "20-why-does-the-issue-or-challenge-need-to-be-addressed-now": "WhyProblemNeedsAddressed",
    "21-what-has-prevented-the-resolution-of-the-issue-or-challenge-e-g-lack-of-funds-to-purchase-curriculum-materials-teacher-training-software-etc": "ProblemPrevention",
    "22-what-type-of-measurable-data-do-you-have-to-show-there-is-a-need-to-help-your-students": "MeasurableData",
    "23-what-does-the-quantitative-or-measurable-data-reveal-about-the-issue-or-challenge-concerning-student-learning-in-your-classroom": "QuantitativeData",
    "24-with-support-from-a-grant-funder-what-major-steps-will-you-propose-to-address-the-issue-or-challenge-and-reduce-the-learning-gap": "ProposedIdea",
    "25-why-is-your-project-idea-for-change-significant-unique-and-or-feasible": "SupportReason",
    "26-what-evidence-based-practices-or-research-generally-support-your-project-idea-for-now-you-do-not-need-to-provide-exact-sources-just-provide-a-summary": "ResearchIdea",
    "27-what-is-the-primary-goal-for-your-proposed-project-it-should-be-a-broad-and-positive-statement-of-what-you-aim-to-accomplish-an-example-of-a-goal-is-implement-the-xyz-literacy-tutoring": "PrimaryGoal",
    "28a-under-objective-1-of-my-project-i-will": "ObjectiveOne",
    "28b-under-objective-2-of-my-project-i-will": "ObjectiveTwo",
    "28c-under-objective-3-of-my-projcet-i-will": "ObjectiveThree",
    "29-what-category-does-your-project-address": "ProjectCategory",
    "30-what-is-the-name-or-title-of-your-project": "ProjectTitle",
    "31-how-many-students-will-be-served": "StudentsServed",
    "32-what-grades-does-your-project-address": "GradesAddressed",
    "33-in-accordance-with-your-goal-and-objectives-what-do-you-want-to-accomplish-with-your-students-emphasize-innovative-activities-that-will-improve-student-learning": "AccomplishmentGoals",
    "34a-action-step-1": "ActionStep1",
    "34b-action-step-2": "ActionStep2",
    "34c-action-step-3": "ActionStep3",
    "34d-action-step-4": "ActionStep4",
    "34e-action-step-5": "ActionStep5",
    "35a-what-is-the-name-of-the-program-model-or-initiative-that-your-project-based-on": "BasedOnResearchName",
    "35b-what-is-the-organizations-name-that-created-the-model-initiative-or-concept": "PickedProjectOrganization",
    "35c-why-did-you-pick-the-model-initiative-or-concept-as-the-basis-for-your-project": "WhyPickedProject",
    "36-what-makes-your-project-approach-significant-or-innovative": "ProjectSignificance",
    "37-how-will-your-students-be-identified-to-participate-in-your-project": "IdentifyStudents",
    "38a-standard-1": "StandardOne",
    "38b-standard-2": "StandardTwo",
    "38c-standard-3": "StandardThree",
    "39a-by-participating-in-my-project-students-will-reach-the-following-first-learning-outcome": "ExampleOne",
    "39b-by-participating-in-my-project-students-will-reach-the-following-second-learning-outcome": "ExampleTwo",
    "39c-by-participating-in-my-project-students-will-reach-the-following-third-learning-outcome": "ExampleThree",
    "40-in-which-subjects-will-your-project-promote-student-achievement": "PickedProjectName",
    "41-what-curriculum-will-you-use-to-measure-changes-in-the-achievement-of-students-participating-in-your-project": "CurriculumMeasurement",
    "42-which-assessments-will-you-use-to-measure-student-achievement-for-those-participating-in-your-project": "MeasureStudentAchievement",
    "43-why-have-you-selected-this-assessment": "WhyThisAssessment",
    "44-what-have-previous-assessment-results-revealed-to-you-about-the-achievement-of-your-students-which-prompted-you-to-pursue-grant-funding": "AssessmentResults",
    "45-by-what-percentage-do-you-expect-student-achievement-to-increase-due-to-your-project": "PercentageIncrease",
    "46-how-many-months-will-it-take-to-provide-a-final-assessment-of-your-students-achievement-under-your-project": "FinalAssessmentMonths",
    "47-what-research-or-evidence-based-practices-is-your-project-based-on": "ProjectApproachResearch",
    "48-what-do-research-or-evidence-based-practices-say-to-support-the-success-of-your-project-idea": "RequiredProfessionalDevelopoment",
    "50-describe-the-professional-development-you-need-to-make-your-project-successful-e-g-coaching-one-on-one-mentoring-etc": "TeacherPractices",
    "51-is-your-approach-to-professional-development-supported-by-a-framework-that-helps-you-teachers-improve-their-learning": "ProfessionalDevelopomentBasis",
    "52-does-your-professional-development-incorporate-data-analysis": "ProfessionalDevelopomentData",
    "53-how-will-the-professional-development-change-your-teaching-practices-in-the-classroom": "TeacherPractices",
    "54a-when-do-you-expect-to-start-your-project": "ProjectStart",
    "54b-when-do-you-expect-to-end-your-project": "ProjectEnd",
    "55-what-will-you-do-during-the-first-three-months-of-your-project": "FirstThreeMonths",
    "56-what-will-you-do-during-months-4-5-of-your-project": "Months4_5",
    "57-what-will-you-do-during-months-6-9-of-your-project": "Months6_9",
    "58-what-will-you-do-during-months-10-12-of-your-project": "Months10_12",
    "59-who-will-be-the-lead-person-to-manage-your-project": "LeadManager",
    "60-what-is-the-title-of-the-person-who-will-manage-your-project": "LeadManagerTitle",
    "61-who-will-your-project-leader-report-to": "ManagerReportsTo",
    "62a-name": "Person1Name",
    "62a1-title": "Person1Title",
    "62b-name": "Person2Name",
    "62b1-title": "Person2Title",
    "62c-name": "Person3Name",
    "62c1-title": "Person3Title",
    "62d-name": "Person4Name",
    "62d1-title": "Person4Title",
    "63a-name": "Person5Name",
    "63b-title": "Person5Title",
    "64-how-often-will-your-leadership-team-meet-to-review-the-progress-of-your-project": "ReviewProgressMeeting",
    "65-how-much-grant-funding-do-you-need-for-your-project": "GrantAmount",
    "66-how-do-you-plan-to-use-the-grant-funding-for-your-project": "UsingFunds",
    "67-are-there-other-resources-or-funds-committed-to-this-project": "OtherResources"
}

def escape_xml(text):
    """Escape XML special characters"""
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#x2019;')

def normalize_xml_for_placeholders(xml_content):
    """
    Remove Word's XML formatting that splits placeholders across multiple <w:t> tags.
    Word often splits text when there's formatting, which breaks our placeholder replacement.
    """
    # Remove text run boundaries that might split placeholders
    # This merges adjacent <w:t> tags within the same run
    xml_content = re.sub(r'</w:t>(\s*)<w:t[^>]*>', r'\1', xml_content)
    
    # Also handle text runs with properties in between
    xml_content = re.sub(r'</w:t></w:r><w:r[^>]*><w:t[^>]*>', '', xml_content)
    
    return xml_content

def process_docx(template_data, form_data):
    """Fill Word template with form data"""
    # Open template
    template_zip = zipfile.ZipFile(io.BytesIO(template_data))
    doc_xml = template_zip.read('word/document.xml').decode('utf-8')
    
    print("=== BEFORE NORMALIZATION ===")
    print(f"Template length: {len(doc_xml)} chars")
    sample_placeholder_count = doc_xml.count('${SchoolName}')
    print(f"Found ${{SchoolName}} before normalization: {sample_placeholder_count}")
    
    # Normalize XML to merge split placeholders
    doc_xml = normalize_xml_for_placeholders(doc_xml)
    
    print("=== AFTER NORMALIZATION ===")
    print(f"Template length: {len(doc_xml)} chars")
    sample_placeholder_count = doc_xml.count('${SchoolName}')
    print(f"Found ${{SchoolName}} after normalization: {sample_placeholder_count}")
    
    # Count total placeholders
    total_placeholders_found = 0
    for word_placeholder in FIELD_MAPPING.values():
        placeholder = '${' + word_placeholder + '}'
        count = doc_xml.count(placeholder)
        total_placeholders_found += count
    
    print(f"Total placeholders found in template: {total_placeholders_found}")
    
    # DEBUG: Log form data
    print("=== FORM DATA ===")
    print(f"Number of fields received: {len(form_data)}")
    non_empty_fields = sum(1 for v in form_data.values() if v)
    print(f"Non-empty fields: {non_empty_fields}")
    
    replacements_made = 0
    warnings = 0
    
    # Replace all placeholders
    for sureforms_field, word_placeholder in FIELD_MAPPING.items():
        value = form_data.get(sureforms_field, '')
        placeholder = '${' + word_placeholder + '}'
        
        if value:
            before_count = doc_xml.count(placeholder)
            doc_xml = doc_xml.replace(placeholder, escape_xml(value))
            after_count = doc_xml.count(placeholder)
            
            if before_count > 0:
                print(f"✓ Replaced {before_count}x '{placeholder}' with '{value[:30]}...'")
                replacements_made += before_count
            else:
                warnings += 1
                if warnings <= 5:  # Only show first 5 warnings to avoid spam
                    print(f"⚠ '{placeholder}' not found (have value: '{value[:20]}...')")
    
    print(f"=== SUMMARY ===")
    print(f"Total replacements made: {replacements_made}")
    print(f"Total warnings (placeholder not found): {warnings}")
    
    # Create new docx
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as new_docx:
        for item in template_zip.namelist():
            if item != 'word/document.xml':
                new_docx.writestr(item, template_zip.read(item))
        new_docx.writestr('word/document.xml', doc_xml.encode('utf-8'))
    
    output.seek(0)
    return output

@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.utcnow().isoformat(),
        'template_configured': 'TEMPLATE_BASE64' in os.environ
    })

@app.route('/fill-document', methods=['POST'])
def fill_document():
    """Main endpoint - receives Sure Forms data, returns filled document"""
    try:
        # Get form data
        data = request.json
        if isinstance(data, list) and len(data) > 0:
            form_data = data[0].get('body', {})
        elif 'body' in data:
            form_data = data['body']
        else:
            form_data = data
        
        # Get template from environment variable (base64 encoded)
        template_b64 = os.getenv('TEMPLATE_BASE64')
        if not template_b64:
            return jsonify({'error': 'Template not configured. Set TEMPLATE_BASE64 environment variable.'}), 500
        
        template_data = base64.b64decode(template_b64)
        
        # Process document
        filled_doc = process_docx(template_data, form_data)
        
        # Create filename
        school_name = form_data.get('1-what-is-the-name-of-your-school', 'School')
        filename = f"Grant_Application_{school_name.replace(' ', '_')}.docx"
        
        # Return filled document
        return send_file(
            filled_doc,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        print(f"=== ERROR ===")
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        print(f"=== END ERROR ===")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
