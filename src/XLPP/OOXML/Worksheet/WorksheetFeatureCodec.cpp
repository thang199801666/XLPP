#include "OOXML/Worksheet/WorksheetFeatureCodec.h"
#include "Package/Xml/XmlUtilities.h"

using xlpp::internal::xmlEscape;

namespace xlpp::internal::ooxml {

std::string filterOperatorName(xlpp::FilterOperator op) {
    switch (op) {
    case xlpp::FilterOperator::NotEqual: return "notEqual";
    case xlpp::FilterOperator::LessThan: return "lessThan";
    case xlpp::FilterOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::FilterOperator::GreaterThan: return "greaterThan";
    case xlpp::FilterOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    default: return "equal";
    }
}

xlpp::FilterOperator parseFilterOperator(const std::string& value) {
    if (value == "notEqual") return xlpp::FilterOperator::NotEqual;
    if (value == "lessThan") return xlpp::FilterOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::FilterOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::FilterOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::FilterOperator::GreaterThanOrEqual;
    return xlpp::FilterOperator::Equal;
}

std::string conditionalOperatorName(xlpp::ConditionalOperator op) {
    switch (op) {
    case xlpp::ConditionalOperator::NotEqual: return "notEqual";
    case xlpp::ConditionalOperator::LessThan: return "lessThan";
    case xlpp::ConditionalOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::ConditionalOperator::GreaterThan: return "greaterThan";
    case xlpp::ConditionalOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    case xlpp::ConditionalOperator::Between: return "between";
    case xlpp::ConditionalOperator::NotBetween: return "notBetween";
    default: return "equal";
    }
}

xlpp::ConditionalOperator parseConditionalOperator(const std::string& value) {
    if (value == "notEqual") return xlpp::ConditionalOperator::NotEqual;
    if (value == "lessThan") return xlpp::ConditionalOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::ConditionalOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::ConditionalOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::ConditionalOperator::GreaterThanOrEqual;
    if (value == "between") return xlpp::ConditionalOperator::Between;
    if (value == "notBetween") return xlpp::ConditionalOperator::NotBetween;
    return xlpp::ConditionalOperator::Equal;
}

// Write a <cfvo> element for data bars / color scales / icon sets.
// Formula thresholds are represented by the cfvo val attribute; a sibling
// <f> element is not part of CT_Cfvo and is rejected by strict consumers.
void writeCfvo(std::ostringstream& xml, const xlpp::Cfvo& cfvo) {
    xml << "<cfvo type=\"" << xmlEscape(cfvo.type) << "\"";
    if (!cfvo.formula.empty())
        xml << " val=\"" << xmlEscape(cfvo.formula) << "\"";
    else if (cfvo.hasValue)
        xml << " val=\"" << cfvo.value << "\"";
    xml << "/>";
}

xlpp::Cfvo parseCfvo(const std::string& tag) {
    xlpp::Cfvo result;
    result.type = xlpp::internal::attribute(tag, "type");
    const auto value = xlpp::internal::attribute(tag, "val");
    if (!value.empty()) {
        if (result.type == "formula") {
            result.formula = value;
            result.hasValue = true;
        } else {
            result.value = std::stod(value);
            result.hasValue = true;
        }
    }
    return result;
}


std::string dataValidationTypeName(xlpp::DataValidationType type) {
    switch (type) {
    case xlpp::DataValidationType::Whole: return "whole";
    case xlpp::DataValidationType::Decimal: return "decimal";
    case xlpp::DataValidationType::List: return "list";
    case xlpp::DataValidationType::Date: return "date";
    case xlpp::DataValidationType::Time: return "time";
    case xlpp::DataValidationType::TextLength: return "textLength";
    case xlpp::DataValidationType::Custom: return "custom";
    default: return "none";
    }
}
xlpp::DataValidationType parseDataValidationType(const std::string& value) {
    if (value == "whole") return xlpp::DataValidationType::Whole;
    if (value == "decimal") return xlpp::DataValidationType::Decimal;
    if (value == "list") return xlpp::DataValidationType::List;
    if (value == "date") return xlpp::DataValidationType::Date;
    if (value == "time") return xlpp::DataValidationType::Time;
    if (value == "textLength") return xlpp::DataValidationType::TextLength;
    if (value == "custom") return xlpp::DataValidationType::Custom;
    return xlpp::DataValidationType::None;
}
std::string dataValidationOperatorName(xlpp::DataValidationOperator op) {
    switch (op) {
    case xlpp::DataValidationOperator::NotBetween: return "notBetween";
    case xlpp::DataValidationOperator::Equal: return "equal";
    case xlpp::DataValidationOperator::NotEqual: return "notEqual";
    case xlpp::DataValidationOperator::LessThan: return "lessThan";
    case xlpp::DataValidationOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::DataValidationOperator::GreaterThan: return "greaterThan";
    case xlpp::DataValidationOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    default: return "between";
    }
}
xlpp::DataValidationOperator parseDataValidationOperator(const std::string& value) {
    if (value == "notBetween") return xlpp::DataValidationOperator::NotBetween;
    if (value == "equal") return xlpp::DataValidationOperator::Equal;
    if (value == "notEqual") return xlpp::DataValidationOperator::NotEqual;
    if (value == "lessThan") return xlpp::DataValidationOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::DataValidationOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::DataValidationOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::DataValidationOperator::GreaterThanOrEqual;
    return xlpp::DataValidationOperator::Between;
}
std::string dataValidationErrorStyleName(xlpp::DataValidationErrorStyle style) {
    switch (style) {
    case xlpp::DataValidationErrorStyle::Warning: return "warning";
    case xlpp::DataValidationErrorStyle::Information: return "information";
    default: return "stop";
    }
}
xlpp::DataValidationErrorStyle parseDataValidationErrorStyle(const std::string& value) {
    if (value == "warning") return xlpp::DataValidationErrorStyle::Warning;
    if (value == "information") return xlpp::DataValidationErrorStyle::Information;
    return xlpp::DataValidationErrorStyle::Stop;
}


} // namespace xlpp::internal::ooxml
