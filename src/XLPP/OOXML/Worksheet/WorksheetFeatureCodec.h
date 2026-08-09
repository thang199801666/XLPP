#pragma once

#include <XLPP/Worksheet/ConditionalFormatting.h>
#include <XLPP/Worksheet/DataValidation.h>
#include <XLPP/Worksheet/Filters.h>

#include <sstream>
#include <string>

namespace xlpp::internal::ooxml {

std::string filterOperatorName(xlpp::FilterOperator op);
xlpp::FilterOperator parseFilterOperator(const std::string& value);
std::string conditionalOperatorName(xlpp::ConditionalOperator op);
xlpp::ConditionalOperator parseConditionalOperator(const std::string& value);
void writeCfvo(std::ostringstream& xml, const xlpp::Cfvo& cfvo);
xlpp::Cfvo parseCfvo(const std::string& tag);
std::string dataValidationTypeName(xlpp::DataValidationType type);
xlpp::DataValidationType parseDataValidationType(const std::string& value);
std::string dataValidationOperatorName(xlpp::DataValidationOperator op);
xlpp::DataValidationOperator parseDataValidationOperator(const std::string& value);
std::string dataValidationErrorStyleName(xlpp::DataValidationErrorStyle style);
xlpp::DataValidationErrorStyle parseDataValidationErrorStyle(const std::string& value);

} // namespace xlpp::internal::ooxml
