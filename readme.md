# Lysna automated analysis
This repo contains code and materials (cleaned of proprietary information) for the automated analysis of the survey-based [Lysna](https://lysna.world/) team assessment.

## The Lysna assessment
The Lysna team assessment is a survey-based assessment of team dynamics, including psychological safety, attachment-informed characteristics, and other group metrics relevant to team culture and performance.

## Why automate it?
In the past, once a team completed the survey, a Lysna consultant would manually analyse the data to extracting key findings from the results, often having to spend several days of work to go in-depth on each subscale of the assessment.
This application solves that problem by automating the preprocessing and statistical analysis of the assessment.

This results in
1. time-savings;
2. standardising the delivered product; and,
3. produces visuals to support the reporting and delivery of key findings.

## What does it do?
To begin the analysis, the app automatically detects whether the assessment respondents belong to the same singular team or to several teams within one organisation.
If there are several teams, they will be analysed both seperately and as one, producing $n+1$ result outputs.

Within each subscale, the application identifies the
- Highest scoring items
- Lowest scoring items
- Items with greatest variance (signalling disagreement within the team)
- Items with lowest variance (signalling agreement)

To support the delivery of results to clients, the app
- Plots each item
- Plots each scale
- Has an option for dynamically adding colour to plots to represent scores above or below normative values.
