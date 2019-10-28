import sys
import pandas as pd
import re
import markdown
import numpy as np
import pypandoc
import numbers
import codecs
import os
import math
from time import gmtime, strftime

class Evaluator():
    def __init__(self, filename, str_course_code, str_course_semester):
        """ Evaluator instances must be given a filename from which data are to be read. """
        self.filename = filename
        self.xls_file = pd.ExcelFile(filename)
        self.ncols_criteria = 9 # how many columns in each activity is to be treated as evaluation criteria?
        self.str_activity = "Aktivitet"
        self.str_rubric = "vurderingsrubrikk"
        self.str_report = "Vurderingsrapport"
        self.str_criterion_type = "Kriterietype"
        self.str_criterion_theme = "Vurderingskriterium"
        self.str_course_code = str_course_code
        self.str_course_semester = str_course_semester
        self.str_commentsrow = 'Spesifikke kommentarer til hvert punkt'
        self.str_category = 'Kategori'
        self.sheet_names = pd.ExcelFile(filename).sheet_names
        self.rubric_sheetnames = self.find_evaluation_rubric_sheetnames()
        self.rubric_sheetidxs = self.find_evaluation_rubric_sheetname_idxs()
        self.activity_numbers = self.find_activity_numbers()
        self.rubrics = self.find_rubrics()
        self.comments_rows = self.find_comments_rows()
        self.comments = self.find_comments()
        self.points = self.find_points()
        self.criteria_colnames = [self.find_criteria_colnames(activity_number) for activity_number in self.activity_numbers]
        self.achievement_level_low = "Lav måloppnåelse"
        self.achievement_level_mid = "Middels måloppnåelse"
        self.achievement_level_hi = "Høy måloppnåelse"
        self.str_achievement = "Måloppnåelse"
        self.str_combine_achievement_levels_single = ""#"Score: "#"På dette vurderingskriteriet oppnår du "
        self.str_combine_achievement_levels_start = ""#"Score"#"På dette vurderingskriteriet oppnår du "
        self.str_combine_achievement_levels_mid = "-"
        self.str_tempdir = "tmp"
        self.pandoc_args = ['--mathjax', '-V', 'geometry:margin=2.5cm']
        self.str_summary = "Oppsummering"
        self.str_feedback = "Detaljert tilbakemelding"
        self.str_reason_for_lowerscore = "Avvik fra høyeste måloppnåelse og/eller andre kommentarer"
        self.str_points = "poeng"
        self.str_table_score = "Score"
        self.str_table_normalised_score = "Vektet score"
        self.str_intermediate_performance = "På dette vurderingskriteret har du prestert et sted mellom følgende måloppnåelsebeskrivelser:"
        self.str_performance = "Følgende måloppnåelsebeskrivelse er omtrent beskrivende for prestasjonen din på dette vurderingskriteriet:"


    def read_evaluation_rubrics(self, ):
        """ Reads an evaluation rubric from an .xlsx file """
        return self.xls_file.parse()
    
    def find_evaluation_rubric_sheetnames(self):
        """ Finds the names of the sheets that contain evaluation rubrics. """
        sn = pd.Series(self.xls_file.sheet_names)
        sheet_idxs_containing_rubrics = sn.str.contains(self.str_activity) & sn.str.contains(self.str_rubric)
        idx_rubrics = [i for i, x in enumerate(sheet_idxs_containing_rubrics) if x]
        sheetnames_rubrics = sn[idx_rubrics]
        return [x for x in sheetnames_rubrics.values]
    
    def find_activity_numbers(self):
        activity_numbers = [int(re.search("\d+", sheetname).group()) for sheetname in self.rubric_sheetnames]
        
        if not len(set(activity_numbers)) == len(activity_numbers):
            raise AssertionError("Duplicate activity numbers were found. Please fix.")
        return activity_numbers
    
    def find_evaluation_rubric_sheetname_idxs(self):
        """ Finds the indices of the sheetnames that contain evaluation rubrics. """
        sn = pd.Series(self.xls_file.sheet_names)
        sheet_idxs_containing_rubrics = sn.str.contains(self.str_activity) & sn.str.contains(self.str_rubric)
        sheetname_idxs = [i for i, x in enumerate(sheet_idxs_containing_rubrics) if x]
        return sheetname_idxs
    
    def find_rubrics(self, skiprows = 2):
        """ Finds and parses the rubric sheets of the rubric excel files as individual pandas dataframes."""
        return [self.xls_file.parse(sheetname, skiprows = skiprows) for sheetname in self.rubric_sheetnames]
        
    def validate_activity_number(self, activity_number):
        if not activity_number in self.activity_numbers:
            err_msg = "Activity number " + str(activity_number) + \
                " doesn't exist. Activity numbers that do exist are " + \
                str(self.activity_numbers)
            raise AssertionError(err_msg)
            
    
    def find_comments_row(self, activity_number, skiprows = 2, s_category = 'Kategori', s_row = 'Spesifikke kommentarer til hvert punkt'):
        """
        Finds the row at which the comments start in a sheet containing an evaluation 
        rubric. The first rows in such a sheet are just the points given. Those rows 
        are followed by an identical copy of the rows where the points are replaced 
        with comments to the specific criteria. We need to find that row, so that we 
        can extract the comments corresponding to the points.
        """
        self.validate_activity_number(activity_number)
            
        # Find the row index of the row just before the comments section starts
        rubric_idx = self.activity_numbers.index(activity_number)
        
        df = self.rubrics[rubric_idx]
        matches_str_category = df[self.str_category].values == self.str_commentsrow
        idx_comments = np.where(matches_str_category)[0][0] + skiprows
        return idx_comments
                              
    def find_criteria_colnames(self, activity_number):
        df =  self.get_points(activity_number, include_evaluation_criteria = True)
        return list(df.columns[0:self.ncols_criteria].values)
    
    def get_students(self, activity_number):
        activity_idx = self.activity_numbers.index(activity_number)
        return self.comments[activity_idx].columns[(self.ncols_criteria):]
            
    def get_categories(self, activity_number):
        activity_idx = self.activity_numbers.index(activity_number)

        return list(self.comments[activity_idx][self.str_category].unique())
    
    def get_category_counts(self, activity_number):
        activity_idx = self.activity_numbers.index(activity_number)
        return list(self.comments[activity_idx][self.str_category].value_counts())
    
    def find_comments_rows(self):
        return [self.find_comments_row(activity_number) for activity_number in self.activity_numbers]
    
    def find_comments(self):
        comments = [self.rubrics[i].iloc[self.comments_rows[i]:]  for i in range(len(self.rubrics))]
        
        comments_sorted = [c.sort_values(by = [self.str_category, self.str_criterion_type, self.str_criterion_theme], 
                                        ascending = [1, 1, 1]) for c in comments]
         

        return comments_sorted
    
    def find_points(self,  nrows_separating_comments = 2):
        pts = [self.rubrics[i].iloc[:(self.comments_rows[i] - (nrows_separating_comments + 1))] for i in range(len(self.rubrics))]
        
        pts_sorted = [p.sort_values(by = [self.str_category, self.str_criterion_type, self.str_criterion_theme], 
                                   ascending = [1, 1, 1]) for p in pts]
         
        return pts_sorted
        
    
    def get_color_from_score(self, score):
        """ Color according to the score """
        if score < 1:
            return "Brown"#"brown"
        elif score == 1: 
            return "Red"#"red"
        elif score < 2:
            return "RedOrange"
        elif score == 2:
            return "Orange"
        elif score < 3:
            return "YellowOrange"
        elif score == 3:
            return "Green"
        elif score > 3:
            return "Aquamarine"
    
    def get_comments(self, activity_number, include_evaluation_criteria = False):
        """ Get the comments for a specific activity, specified by its activity number. """
        self.validate_activity_number(activity_number)
        df = self.comments[self.activity_numbers.index(activity_number)]
        df = df.sort_values(by = [self.str_category, self.str_criterion_type, self.str_criterion_theme], ascending = [1, 1, 1])

        if include_evaluation_criteria:
            return df
        else:
            return df[df.columns[self.ncols_criteria:]]
    
    def get_student_comments(self, student, activity_number, include_evaluation_criteria = False):
        """ 
        Get the points a student got on a particular activity. Evaluation criteria can be included together with 
        the points if necessary.
        """
        if not include_evaluation_criteria:
            df_comments = self.get_comments(activity_number, include_evaluation_criteria = False)
            return df_comments[student]
        else:
            df_comments = self.get_comments(activity_number, include_evaluation_criteria = True)
            cols = self.find_criteria_colnames(activity_number)
            cols.append(student)
            return df_comments[cols]

    
    def get_points(self, activity_number, include_evaluation_criteria = False):
        """ Get the comments for a specific activity, specified by its activity number, 
            as a pandas dataframe. Optionally, the evaluation criteria can be included.
        """
        self.validate_activity_number(activity_number)
        df = self.points[self.activity_numbers.index(activity_number)]
        df = df.sort_values(by = [self.str_category, self.str_criterion_type, self.str_criterion_theme], ascending = [1, 1, 1])

        if include_evaluation_criteria:
            return df
        else:
            return df[df.columns[self.ncols_criteria:]]
        
    
    def get_student_points(self, student, activity_number, include_evaluation_criteria = False):
        """ 
        Get the points a student got on a particular activity. Evaluation criteria can be included together with 
        the points if necessary.
        """
        if not include_evaluation_criteria:
            df_points = self.get_points(activity_number, include_evaluation_criteria = False)
            return df_points[student]
        else:
            df_points = self.get_points(activity_number, include_evaluation_criteria = True)
            cols = self.find_criteria_colnames(activity_number)
            cols.append(student)
            return df_points[cols]

    
    def get_achievement_level(self, score, normalised = False):
        """ 
        Gets the achievement level(s) based on a score. If normalised = False, 
        then it is assumed that scores are given either as 1, 2 or 3, corresponding
        to "low", "mid" and "high" achievement level. If scores are between those 
        values, multiple achievement levels will be returned.
        
        Returns the achievement level as a text string that can be used to index a 
        score or comment dataframe located in self.points or self.comments.
        """
        if not normalised:
            if score <= 1:
                return self.achievement_level_low
            elif score < 2:
                return [self.achievement_level_low, self.achievement_level_mid]
            elif score == 2:
                return self.achievement_level_mid
            elif score < 3:
                return [self.achievement_level_mid, self.achievement_level_hi]
            elif score == 3:
                return self.achievement_level_hi
    
    def get_student_achievement_levels(self, student, activity_number, normalised = False):
        if not normalised:
            score = self.get_student_points(student, activity_number, include_evaluation_criteria = True)
            
            achievement_levels = score[score.columns[-1]].apply(self.get_achievement_level)
            
            return [x for x in achievement_levels.values]
    
    def get_combined_achievement_levels(self, student, activity_number, 
                                        normalised = False, colors = False):
        
        achievement_levels = self.get_student_achievement_levels(student, activity_number)
        scores = self.get_student_points(student, activity_number).values
        
        combined_achievement_levels = list()
        
        for i, level in enumerate(achievement_levels):
            if isinstance(level, list):
                s = self.str_combine_achievement_levels_start + level[0].split()[0].lower() + \
                    self.str_combine_achievement_levels_mid + level[1].lower()

            else:
                s = self.str_combine_achievement_levels_single + level.lower()
            
            if colors:
                s = "\\textcolor{"+ self.get_color_from_score(scores[i]) +"}{" + s + "}"
            combined_achievement_levels.append(s)

        return combined_achievement_levels
    
    
    def get_generic_comments(self, achievement_levels, activity_number):
        """ Given a list of achievement levels, generate generic comments. """
        activity_idx = self.activity_numbers.index(activity_number)
        df_scores = self.points[activity_idx]
        
        comments = list()
        for i, level in enumerate(achievement_levels):
            comment = df_scores[level].iloc[i]
            
            # Depending on the score, there might be several comments applicable, so get 
            # all of them.
            if isinstance(comment, pd.Series):
                comments.append([x for x in comment.values])
            else:
                comments.append(comment)
        return comments
    
    
    def get_combined_generic_comments(self, student, activity_number, colors = True):
        scores = self.get_student_points(student, activity_number).values
        achievement_levels = self.get_student_achievement_levels(student, activity_number)
        comments = self.get_generic_comments(achievement_levels, activity_number)
        
        combined_generic_comments = list()
        for i, (level, comment, score) in enumerate(zip(achievement_levels, comments, scores)):
            #print(i, level, comment, score, "\n")
            
            # If level is a list of achievement levels, combine the generic comments at both levels
            if isinstance(level, list):

                col_lo = self.get_color_from_score(math.floor(score))
                col_hi = self.get_color_from_score(math.ceil(score))

                comb_comments = "".join([
                    "".join([self.str_intermediate_performance, "\n\n"]),
                    "".join(["> *" + comment[0] + "*\n\n"]),
                    "".join(["> *" + comment[1] + "*\n\n"])
                ])
                combined_generic_comments.append(comb_comments)
            else:
                gen_comment = "" 
                gen_comment += self.str_performance + "\n\n"
                gen_comment += "> *" + comment + "*\n\n"
                
                combined_generic_comments.append(gen_comment)
                
                
        return combined_generic_comments
    
    def get_criteria(self, category, activity_number):
        activity_idx = self.activity_numbers.index(activity_number)
        df = self.comments[activity_idx]
        criteria = df[df[self.str_category] == category][self.str_criterion_theme]
        return list(criteria[~criteria.isnull()].values)
    
    def get_complete_criteria(self, activity_number):
        activity_idx = self.activity_numbers.index(activity_number)
        df = self.comments[activity_idx]
        return df
        criteria = df[[self.str_category, self.str_criterion_theme]]
        return list(criteria[~criteria.isnull()].values)
           
    def write_report_to_file(self, str_report, student, activity_number, 
                             temp = False, timestamp = True, 
                             remove_temp_files = True, toc = False):
        
        path = ""
        if temp:
            path += "./" + self.str_tempdir + "/"
        else:
            path += "./"
        
        path += self.str_course_code + "/" + \
                self.str_course_semester + "/" + self.str_activity + "_" + \
                str(activity_number) + "/" + student + "/"

        os.makedirs(os.path.dirname(path), exist_ok=True)
        fn_student = path + \
            self.str_report + "_" + self.str_course_code + "_" + \
            self.str_course_semester + "_" + self.str_activity + "_" + str(activity_number) + "_" + \
            student.replace(" ", "_").replace(",", "") #+ t
        
        if temp:
            fn_student += "_tmp"
        
        if timestamp: 
            fn_student += strftime("_%Y%m%d_%H%M%S", gmtime())

        fn_txt = fn_student + ".txt"
        fn_md = fn_student + ".md"
        fn_pdf = fn_student + ".pdf"
        fn_html = fn_student + ".html"
            
        os.makedirs(os.path.dirname(fn_md), exist_ok=True)
        with open(fn_md, "w+") as f:
            f.write(str_report)
            f.close()
                    
        # Write report as text first
        os.makedirs(os.path.dirname(fn_txt), exist_ok=True)
        with open(fn_txt, "w+") as f:
            f.write(str_report)
            f.close()

        # Then read text file and create html
        input_file = codecs.open(fn_txt, mode="r", encoding="utf-8")
        text = input_file.read()
        html = markdown.markdown(text)
                
        # Write html file
        os.makedirs(os.path.dirname(fn_html), exist_ok=True)
        output_file = codecs.open(fn_html, "w", encoding="latin-1",errors="xmlcharrefreplace")
        output_file.write(html)

        pargs = self.pandoc_args
            
        pargs.append('--highlight-style=pygments')
        pargs.append('--include-in-header=config/header.tex')
        pargs.append('--include-after-body=config/after-body.tex')
        pargs.append('--pdf-engine=xelatex')
        
        if toc: 
            pargs.append('--table-of-contents')
            
        #print(pargs)
        # Convert temporary markdown file to pdf using pandoc
        output = pypandoc.convert_file(fn_md, 'pdf', format="markdown",
                    outputfile = fn_pdf, 
                    extra_args = pargs)
        
        if remove_temp_files:
            os.remove(fn_txt)
            os.remove(fn_md)
            os.remove(fn_html)
    
    def generate_report(self, student, activity_number, 
                        summary_table = True,
                        colors = True, 
                        export = True, 
                        temp = False, 
                        timestamp = True, 
                        remove_temp_files = True, 
                        toc = False,
                        include_scores = False):
        
        """ Write report as pdf. Requires latex packages textcolorx, environ and tcolorbox, trimspaces. """
        # String that will hold the report.            
        str_report = ''.join(["# ", "Vurdering i ", self.str_course_code, ", ", self.str_course_semester, \
                     " aktivitet ", str(activity_number), " (", student, ")\n\n"])
        
        if summary_table:
            str_report += "## " + self.str_summary + "\n\n"
            str_report += self.make_achievement_level_table(student, activity_number, 
                                                            include_scores = include_scores, 
                                                            colors = colors) + "\n\n"
            str_report += "\\pagebreak"
        #str_report += "## " + self.str_feedback + "\n\n"
        
        
        if include_scores:
            scores = self.get_student_points(student, activity_number).values
        
        # We also need the points to color properly
        student_points = self.get_student_points(student, activity_number).values
        
        for score in student_points:
            if not isinstance(2, numbers.Number):
                any_bad_vals = True
                print(f"REPORT GENERATION FAILED: Missing or incomplete scores for {student}.")
                    
        combined_achievement_levels = self.get_combined_achievement_levels(student, activity_number)
        combined_generic_comments = self.get_combined_generic_comments(student, activity_number)
        specific_comments = self.get_student_comments(student, activity_number, include_evaluation_criteria = False).values
        
        # Go over the criteria in each category
        criteria_counter = 0
        categories = self.get_categories(activity_number)
        categories_counts = self.get_category_counts(activity_number)
            
        total_criteria_counter = 0
        for i, (category, count) in enumerate(zip(categories, categories_counts)):
            # Remember space after # if interpreting as heading
            # Also need double line shift to separate heading from
            # content under that heading.
            criteria_this_category = self.get_criteria(category, activity_number) 
            
            #print("Category: %s" % (category))
            #comments_df = self.get_student_comments(student, activity_number, include_evaluation_criteria = True)
            #scores_df = self.get_student_points(student, activity_number, include_evaluation_criteria = True)
            #print(scores_df)
            #for cat in criteria_this_category:
                #comment = scores_df[level].iloc[i]
                #print(scores_df[level].iloc[i])
                #row = comments_df[comments_df[self.str_criterion_theme].str.contains(cat)]
                #print(row)
                #print("\t%s" % cat)
            
            str_report += '\n# ' + self.str_category + ": " + category + "\n" 
            
            for j, criterion in enumerate(criteria_this_category):
                gen_score = combined_achievement_levels[total_criteria_counter]
                gen_comment = combined_generic_comments[total_criteria_counter]
                specific_comment = specific_comments[total_criteria_counter]
                score = student_points[total_criteria_counter]
                
                # start coloring in header
                if colors:
                    col = self.get_color_from_score(student_points[total_criteria_counter])
                    str_report += '\n## ' + "\\textcolor{" + col + "}{" + criterion
                    
                    # add achievement level in brackets
                    if include_scores:
                        str_report += " [" + gen_score + ", " + str(score) + " " + \
                            self.str_points 
                    else:
                        str_report += " [" + gen_score
                    str_report += "]}" # stop coloring
                else:
                    # add achievement level in brackets
                    str_report += '\n## ' + criterion  + " [" + gen_score + "]"
               
                str_report += "\n\n" + gen_comment
                
                if type(specific_comment) == str:
                    str_report += "**" + self.str_reason_for_lowerscore + "**: "
                    str_report += specific_comment
                    
                str_report += "\n"
                
                # Update the number of criteria we've processed
                total_criteria_counter += 1

            # Another line break for good measure after criteria have been added
            str_report += "\n"
            
        if export:
            try:
                
                self.write_report_to_file(str_report, student, activity_number, 
                    temp = temp, 
                    timestamp = timestamp, 
                    remove_temp_files = remove_temp_files, 
                    toc = toc)
            except Exception as e: 
                print(f"Export failed for {student}")
                print(e)
            
        else:
            return str_report
     
    def make_achievement_level_table(self, student, activity_number, include_scores = False, colors = True):
        df = self.get_complete_criteria(activity_number=activity_number)[['Kategori', 'Vurderingskriterium']]
        categories = df[self.str_category].values
        criteria = df[self.str_criterion_theme].values
        achievement_levels = self.get_combined_achievement_levels(student, activity_number, colors = True)
        scores = self.get_student_points(student, activity_number)
        
        
        if include_scores:
            t = "|" + self.str_category + " | " + self.str_criterion_theme + " | " + self.str_achievement + " | " + self.str_table_score + " | \n"
            t += "|" + "---" + " | " + "---" + " | " + "---" + " | " + "---" + " |\n"

            for i, (cat, crit, lvl, score) in enumerate(zip(categories, criteria, achievement_levels, scores)):
                score_col = self.get_color_from_score(score)
                str_score = "\\textcolor{" + score_col + "}{" + str(score) + "}"
                t += "|" + cat + " | " + crit + " | " + lvl + " | " + str_score + "|\n"

            return t
        else:
            t = "|" + self.str_category + " | " + self.str_criterion_theme + " | " + self.str_achievement + " |\n"
            t += "|" + "---" + " | " + "---" + " | " + "---" + " |\n"

            for i, (cat, crit, lvl) in enumerate(zip(categories, criteria, achievement_levels)):
                t += "|" + cat + " | " + crit + " | " + lvl + " |\n"

            return t
    
    def generate_reports(self, activity_number, **kwargs):
        """ Generate reports for all students for a given activity. """
        students = self.get_students(activity_number)
        
        failed = []
        for student in students:
            scores = self.get_student_points(student, activity_number)
            
            # Test inputs for NaNs or other non-numeric values
            any_bad_vals = False
            if any([isinstance(score, numbers.Number) == False for score in scores]) or any([np.isnan(score) for score in scores]):
                    any_bad_vals = True
                    print(f"REPORT GENERATION FAILED: Missing or incomplete scores for {student}.")
                    failed.append(student)
                    
            else:
                self.generate_report(student, activity_number, **kwargs)
                print(f"SUCCESS: generated report for {student}")
        
        print("Reports could not be generated for the following students:")
        print(failed)


class EvaluationReport():
    def __init__(self, data, data_path, str_report, student, activity_number):
        self.data = data
        self.data_path = data_path
        self.str_report = str_report
        self.student = student 
        self.activity_number = activity_number