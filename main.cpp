#ifndef __PROGTEST__
#include <cstdlib>
#include <cstdio>
#include <cstring>
#include <cctype>
#include <climits>
#include <cfloat>
#include <cassert>
#include <cmath>
#include <iostream>
#include <sstream>
#include <fstream>
#include <iomanip>
#include <string>
#include <array>
#include <vector>
#include <list>
#include <set>
#include <map>
#include <stack>
#include <queue>
#include <unordered_set>
#include <unordered_map>
#include <memory>
#include <algorithm>
#include <functional>
#include <iterator>
#include <stdexcept>
#include <variant>
#include <optional>
#include <compare>
#include <charconv>
#include <span>
#include <utility>
#include "expression.h"
using namespace std::literals;
using CValue = std::variant<std::monostate, double, std::string>;

constexpr unsigned                     SPREADSHEET_CYCLIC_DEPS                 = 0x01;
constexpr unsigned                     SPREADSHEET_FUNCTIONS                   = 0x02;
constexpr unsigned                     SPREADSHEET_FILE_IO                     = 0x04;
constexpr unsigned                     SPREADSHEET_SPEED                       = 0x08;
//constexpr unsigned                     SPREADSHEET_PARSER                      = 0x10;
#endif /* __PROGTEST__ */

// *————————————————————————————————————————————————CPos.h——————————————————————————————————————————————————————* //

/**
 * Class representing position in a spreadsheet
 */
class CPos {
public:
    int m_Row;
    bool m_AbsRow = false;
    int m_Column;
    bool m_AbsColumn = false;

    /**
     * Constructor from string
     * @param str string representation of the position
     */
    explicit CPos(std::string_view str);

    CPos() = default;

    /**
     * Constructor from row and column
     * @param row row
     * @param column column
     */
    CPos(int row, int column) : m_Row(row), m_Column(column) {}

    /**
     * Compare two positions
     * @param rhs right-hand side
     * @return result of the comparison
     */
    std::strong_ordering operator<=>(const CPos &rhs) const;

    /**
     * Save the position to a binary stream
     * @param os output stream
     * @return true if successful
     */
    bool saveBinary(std::ostream &os) const;

    /**
     * Load the position from a binary stream
     * @param is input stream
     * @return true if successful
     */
    bool loadBinary(std::istream &is);

    /**
     * Convert column string to integer
     * @param columnStr column string
     * @return column number
     */
    static int convertColumn(std::string_view columnStr);

};

// *—————————————————————————————————————————————————CPos.cpp——————————————————————————————————————————————————————* //

CPos::CPos(std::string_view str) {
    size_t i = 0;
    size_t len = str.size();

    // Check for absolute column reference
    if (i < len && str[i] == '$') {
        m_AbsColumn = true;
        ++i;
    }

    // Extract column part
    size_t columnStart = i;
    while (i < len && std::isalpha(str[i])) {
        ++i;
    }
    if (i == columnStart) {
        throw std::invalid_argument("Invalid CPos string: missing column part");
    }
    m_Column = convertColumn(str.substr(columnStart, i - columnStart));

    // Check for absolute row reference
    if (i < len && str[i] == '$') {
        m_AbsRow = true;
        ++i;
    }

    // Extract row part
    size_t rowStart = i;
    while (i < len && std::isdigit(str[i])) {
        ++i;
    }
    if (i == rowStart) {
        throw std::invalid_argument("Invalid CPos string: missing row part");
    }
    try {
        m_Row = std::stoi(std::string(str.substr(rowStart, i - rowStart)));
    } catch (const std::exception &) {
        throw std::invalid_argument("Invalid row number");
    }

    if (m_Row < 0) {
        throw std::invalid_argument("Invalid row number: must be positive");
    }

    // Ensure no trailing invalid characters
    if (i != len) {
        throw std::invalid_argument("Invalid CPos string: trailing characters");
    }
}


int CPos::convertColumn(std::string_view columnStr) {
    int column = 0;
    for (char ch: columnStr) {
        if (!std::isalpha(ch)) {
            throw std::invalid_argument("Invalid column part: must be alphabetic");
        }
        column = column * 26 + (std::toupper(ch) - 'A' + 1);
    }
    return column - 1;
}

std::strong_ordering CPos::operator<=>(const CPos &rhs) const {
    return std::tie(m_Row, m_Column) <=> std::tie(rhs.m_Row, rhs.m_Column);
}

bool CPos::saveBinary(std::ostream &os) const {
    // Write row and column as binary integers
    os.write(reinterpret_cast<const char *>(&m_Row), sizeof(m_Row));
    os.write(reinterpret_cast<const char *>(&m_Column), sizeof(m_Column));
    // Write flags as a single byte
    char flags = (m_AbsRow ? 1 : 0) | ((m_AbsColumn ? 1 : 0) << 1);
    os.write(&flags, sizeof(flags));
    return os.good();
}

bool CPos::loadBinary(std::istream &is) {
    is.read(reinterpret_cast<char *>(&m_Row), sizeof(m_Row));
    is.read(reinterpret_cast<char *>(&m_Column), sizeof(m_Column));
    char flags;
    is.read(&flags, sizeof(flags));
    m_AbsRow = flags & 1;
    m_AbsColumn = (flags >> 1) & 1;
    return is.good();
}

// *—————————————————————————————————————————————————COperation.h————————————————————————————————————————————* //
class CCell; // forward declaration

/**
 * abstract class representing an operation in a spreadsheet
 */
class COperation {
public:
    virtual ~COperation() = default;

    /**
     * evaluate the operation
     * @param stack stack of operations
     * @param sheet map of cells
     * @param depth depth of the operation in the stack
     * @return result of the operation
     */
    virtual CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const = 0;

    /**
     * clone the operation
     * @return cloned operation
     */
    virtual std::shared_ptr<COperation> clone() const = 0;

    /**
     * save the operation to a binary stream
     * @param os output stream
     * @return true if successful
     */
    virtual bool saveBinary(std::ostream &os) const = 0;

    /**
     * load the operation from a binary stream
     * @param is input stream
     * @return true if successful
     */
    virtual bool loadBinary(std::istream &is) = 0;

    /**
     * get the type id of the operation
     * @return type id
     */
    virtual int getTypeId() const = 0;

    /**
     * create an operation from a type id
     * @param typeId type id
     * @return shared pointer to the operation
     */
    static std::shared_ptr<COperation> createOperationFromType(int typeId);

};

// *—————————————————————————————————————————————————CAddition.h————————————————————————————————————————————* //

/**
 * class representing addition operation
 */
class CAddition : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CAddition.cpp————————————————————————————————————————————* //

CValue
CAddition::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side))
        return std::get<double>(left_side) + std::get<double>(right_side);

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side))
        return std::get<std::string>(left_side) + std::get<std::string>(right_side);

    return {};
}

std::shared_ptr<COperation> CAddition::clone() const {
    return std::make_shared<CAddition>(*this);
}

bool CAddition::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CAddition::loadBinary(std::istream &is) {
    return is.good();
}

int CAddition::getTypeId() const {
    return 1;
}

// *—————————————————————————————————————————————————CSubtraction.h————————————————————————————————————————————* //

/**
 * class representing subtraction operation
 */
class CSubtraction : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CSubtraction.cpp————————————————————————————————————————————* //

CValue
CSubtraction::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;

    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(left_side) && std::holds_alternative<double>(right_side))
        return std::get<double>(left_side) - std::get<double>(right_side);

    return {};
}

std::shared_ptr<COperation> CSubtraction::clone() const {
    return std::make_shared<CSubtraction>(*this);
}

bool CSubtraction::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CSubtraction::loadBinary(std::istream &is) {
    return is.good();
}

int CSubtraction::getTypeId() const {
    return 2;
}

// *—————————————————————————————————————————————————CMultiplication.h————————————————————————————————————————————* //

/**
 * class representing multiplication operation
 */
class CMultiplication : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CMultiplication.cpp————————————————————————————————————————————* //

CValue CMultiplication::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet,
                                 int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    if (std::holds_alternative<double>(left_side) && std::holds_alternative<double>(right_side))
        return std::get<double>(left_side) * std::get<double>(right_side);

    return {};
}

std::shared_ptr<COperation> CMultiplication::clone() const {
    return std::make_shared<CMultiplication>(*this);
}

bool CMultiplication::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CMultiplication::loadBinary(std::istream &is) {
    return is.good();
}

int CMultiplication::getTypeId() const {
    return 3;
}

// *—————————————————————————————————————————————————CDivision.h————————————————————————————————————————————* //

/**
 * class representing division operation
 */
class CDivision : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CDivision.cpp————————————————————————————————————————————* //

CValue
CDivision::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;

    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(left_side) && std::holds_alternative<double>(right_side)) {
        if (std::get<double>(right_side) == 0)
            return {};
        return std::get<double>(left_side) / std::get<double>(right_side);
    }

    return {};
}

std::shared_ptr<COperation> CDivision::clone() const {
    return std::make_shared<CDivision>(*this);
}

bool CDivision::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CDivision::loadBinary(std::istream &is) {
    return is.good();
}

int CDivision::getTypeId() const {
    return 4;
}

// *—————————————————————————————————————————————————CPower.h————————————————————————————————————————————* //

/**
 * class representing power operation
 */
class CPower : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CPower.cpp————————————————————————————————————————————* //

CValue
CPower::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;

    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(left_side) && std::holds_alternative<double>(right_side)) {
        if (std::get<double>(right_side) == 0)
            return 1.0;
        return std::pow(std::get<double>(left_side), std::get<double>(right_side));

    }

    return {};
}

std::shared_ptr<COperation> CPower::clone() const {
    return std::make_shared<CPower>(*this);
}

bool CPower::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CPower::loadBinary(std::istream &is) {
    return is.good();
}

int CPower::getTypeId() const {
    return 5;
}

// *—————————————————————————————————————————————————CNegation.h————————————————————————————————————————————* //

/**
 * class representing negation operation
 */
class CNegation : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CNegation.cpp————————————————————————————————————————————* //

CValue
CNegation::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;

    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side))
        return -std::get<double>(right_side);

    return {};
}

std::shared_ptr<COperation> CNegation::clone() const {
    return std::make_shared<CNegation>(*this);
}

bool CNegation::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CNegation::loadBinary(std::istream &is) {
    return is.good();
}

int CNegation::getTypeId() const {
    return 6;
}

// *—————————————————————————————————————————————————CEqual.h————————————————————————————————————————————* //

/**
 * class representing equal operation
 */
class CEqual : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CEqual.cpp————————————————————————————————————————————* //

CValue
CEqual::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) == std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) == std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CEqual::clone() const {
    return std::make_shared<CEqual>(*this);
}

bool CEqual::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CEqual::loadBinary(std::istream &is) {
    return is.good();
}

int CEqual::getTypeId() const {
    return 7;
}

// *—————————————————————————————————————————————————CNotEqual.h————————————————————————————————————————————* //

/**
 * class representing not equal operation
 */
class CNotEqual : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CNotEqual.cpp————————————————————————————————————————————* //

CValue
CNotEqual::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) != std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) != std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CNotEqual::clone() const {
    return std::make_shared<CNotEqual>(*this);
}

bool CNotEqual::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CNotEqual::loadBinary(std::istream &is) {
    return is.good();
}

int CNotEqual::getTypeId() const {
    return 8;
}

// *—————————————————————————————————————————————————CLessThan.h————————————————————————————————————————————* //

/**
 * class representing less than operation
 */
class CLessThan : public COperation {

public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CLessThan.cpp————————————————————————————————————————————* //

CValue
CLessThan::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) < std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) < std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CLessThan::clone() const {
    return std::make_shared<CLessThan>(*this);
}

bool CLessThan::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CLessThan::loadBinary(std::istream &is) {
    return is.good();
}

int CLessThan::getTypeId() const {
    return 9;
}

// *—————————————————————————————————————————————————CLessEqual.h————————————————————————————————————————————* //

/**
 * class representing less equal operation
 */
class CLessEqual : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CLessEqual.cpp————————————————————————————————————————————* //

CValue
CLessEqual::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) <= std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) <= std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CLessEqual::clone() const {
    return std::make_shared<CLessEqual>(*this);
}

bool CLessEqual::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CLessEqual::loadBinary(std::istream &is) {
    return is.good();
}

int CLessEqual::getTypeId() const {
    return 10;
}

// *—————————————————————————————————————————————————CGreaterThan.h————————————————————————————————————————————* //

/**
 * class representing greater than operation
 */
class CGreaterThan : public COperation {

public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CGreaterThan.cpp————————————————————————————————————————————* //

CValue
CGreaterThan::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
//    std::cout << "GreaterThan\n";
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) > std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) > std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CGreaterThan::clone() const {
    return std::make_shared<CGreaterThan>(*this);
}

bool CGreaterThan::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CGreaterThan::loadBinary(std::istream &is) {
    return is.good();
}

int CGreaterThan::getTypeId() const {
    return 11;
}

// *—————————————————————————————————————————————————CGreaterEqual.h————————————————————————————————————————————* //

/**
 * class representing greater equal operation
 */
class CGreaterEqual : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CGreaterEqual.cpp————————————————————————————————————————————* //

CValue CGreaterEqual::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet,
                               int &depth) const {
//    std::cout << "GreaterEqual\n";
    depth++;
    CValue right_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);
    CValue left_side = stack[stack.size() - 1 - depth]->evaluate(stack, sheet, depth);

    if (std::holds_alternative<double>(right_side) && std::holds_alternative<double>(left_side)) {
        if (std::get<double>(left_side) >= std::get<double>(right_side))
            return 1.0;
        return 0.0;
    }

    if (std::holds_alternative<std::string>(right_side) && std::holds_alternative<std::string>(left_side)) {
        if (std::get<std::string>(left_side) >= std::get<std::string>(right_side))
            return 1.0;
        return 0.0;
    }

    return {};
}

std::shared_ptr<COperation> CGreaterEqual::clone() const {
    return std::make_shared<CGreaterEqual>(*this);
}

bool CGreaterEqual::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CGreaterEqual::loadBinary(std::istream &is) {
    return is.good();
}

int CGreaterEqual::getTypeId() const {
    return 12;
}

// *—————————————————————————————————————————————————CReference.h——————————————————————————————————————————————————* //

/**
 * class representing a reference to a cell
 */
class CReference : public COperation {
public:
    CReference() = default;

    CReference(std::string &str);

    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    CPos getCPos();

    void setCPos(int rowOffset, int columnOffset);

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;

private:
    CPos m_Pos;
};

// *—————————————————————————————————————————————————CNumber.h——————————————————————————————————————————————————* //

/**
 * class representing a number
 */
class CNumber : public COperation {
public:
    CNumber() = default;

    CNumber(double value);

    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;

private:
    double m_Value;
};

// *—————————————————————————————————————————————————CNumber.cpp——————————————————————————————————————————————————* //

CNumber::CNumber(double value) : m_Value(value) {}

CValue
CNumber::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    return m_Value;
}

std::shared_ptr<COperation> CNumber::clone() const {
    return std::make_shared<CNumber>(*this);
}

bool CNumber::saveBinary(std::ostream &os) const {
    os.write(reinterpret_cast<const char *>(&m_Value), sizeof(m_Value));
    return os.good();
}

bool CNumber::loadBinary(std::istream &is) {
    is.read(reinterpret_cast<char *>(&m_Value), sizeof(m_Value));
    return is.good();
}

int CNumber::getTypeId() const {
    return 14;
}

// *—————————————————————————————————————————————————CString.h——————————————————————————————————————————————————* //

/**
 * class representing a string
 */
class CString : public COperation {
public:
    CString() = default;

    CString(std::string &value);

    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;

private:
    std::string m_Value;
};

// *—————————————————————————————————————————————————CString.cpp——————————————————————————————————————————————————* //

CString::CString(std::string &value) : m_Value(value) {}

CValue
CString::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    depth++;
    return m_Value;
}

std::shared_ptr<COperation> CString::clone() const {
    return std::make_shared<CString>(*this);
}

bool CString::saveBinary(std::ostream &os) const {
    size_t size = m_Value.size();
    os.write(reinterpret_cast<const char *>(&size), sizeof(size));
    os.write(m_Value.data(), size);
    return os.good();
}

bool CString::loadBinary(std::istream &is) {
    size_t size;
    is.read(reinterpret_cast<char *>(&size), sizeof(size));
    m_Value.resize(size);
    is.read(m_Value.data(), size);
    return is.good();
}

int CString::getTypeId() const {
    return 15;
}

// *—————————————————————————————————————————————————CValRange.h——————————————————————————————————————————————————* //

/**
 * class representing a range of values
 */
class CValRange : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CValRange.cpp——————————————————————————————————————————————————* //

CValue
CValRange::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    return {}; // todo
}

std::shared_ptr<COperation> CValRange::clone() const {
    return std::make_shared<CValRange>(*this);
}

bool CValRange::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CValRange::loadBinary(std::istream &is) {
    return is.good();
}

int CValRange::getTypeId() const {
    return 16;
}

// *—————————————————————————————————————————————————CFuncCall.h——————————————————————————————————————————————————* //

/**
 * class representing a function call
 */
class CFuncCall : public COperation {
public:
    CValue
    evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const override;

    std::shared_ptr<COperation> clone() const override;

    bool saveBinary(std::ostream &os) const override;

    bool loadBinary(std::istream &is) override;

    int getTypeId() const override;
};

// *—————————————————————————————————————————————————CFuncCall.cpp——————————————————————————————————————————————————* //

CValue
CFuncCall::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    return {}; // todo
}

std::shared_ptr<COperation> CFuncCall::clone() const {
    return std::make_shared<CFuncCall>(*this);
}

bool CFuncCall::saveBinary(std::ostream &os) const {
    return os.good();
}

bool CFuncCall::loadBinary(std::istream &is) {
    return is.good();
}

int CFuncCall::getTypeId() const {
    return 17;
}

// *—————————————————————————————————————————————————CCell.h——————————————————————————————————————————————————————————————* //

/**
 * Class that represents cell.
 */
class CCell {
public:
    /**
     * Stack of operations.
     */
    std::deque<std::shared_ptr<COperation>> m_Stack;

    /**
     * Flag that indicates whether the cell is calculated.
     */
    bool m_IsCalculated = false;

    CCell() {};

    /**
     * Save cell to binary file.
     * @param os - output stream
     * @return - true if success, false otherwise
     */
    bool saveBinary(std::ostream &os) const;

    /**
     * Calculate cell.
     * @param sheet - sheet of cells
     * @return - result of calculation
     */
    CValue calculateCell(std::map<CPos, CCell> &sheet);

    /**
     * Load cell from binary file.
     * @param is - input stream
     * @return - true if success, false otherwise
     */
    bool loadBinary(std::istream &is);
};

// *—————————————————————————————————————————————————CCell.cpp——————————————————————————————————————————————————————————————* //

CValue CCell::calculateCell(std::map<CPos, CCell> &sheet) {
    if (m_IsCalculated) return {};
    else m_IsCalculated = true;
    int depth = 0;


    auto result = (m_Stack.rbegin()->get()->evaluate(m_Stack, sheet, depth));
    m_IsCalculated = false;
    return result;
}

bool CCell::saveBinary(std::ostream &os) const {
    size_t stackSize = m_Stack.size();
    os.write(reinterpret_cast<const char *>(&stackSize), sizeof(stackSize));

    for (const auto &op: m_Stack) {
        int typeId = op->getTypeId();
        os.write(reinterpret_cast<const char *>(&typeId), sizeof(typeId));
        if (!op->saveBinary(os)) return false;
    }

    bool isCalculated = m_IsCalculated;
    os.write(reinterpret_cast<const char *>(&isCalculated), sizeof(isCalculated));

    return true;
}

// *—————————————————————————————————————————————————CMyExpressionBuilder.h——————————————————————————————————————————————————————* //

class CMyExpressionBuilder : public CExprBuilder {
public:
    /**
     * Add operation to the stack.
     */
    void opAdd() override;

    /**
     * Subtract operation from the stack.
     */
    void opSub() override;

    /**
     * Multiply operation from the stack.
     */
    void opMul() override;

    /**
     * Divide operation from the stack.
     */
    void opDiv() override;

    /**
     * Power operation from the stack.
     */
    void opPow() override;

    /**
     * Negation operation from the stack.
     */
    void opNeg() override;

    /**
     * Equal operation from the stack.
     */
    void opEq() override;

    /**
     * Not equal operation from the stack.
     */
    void opNe() override;

    /**
     * Less than operation from the stack.
     */
    void opLt() override;

    /**
     * Less equal operation from the stack.
     */
    void opLe() override;

    /**
     * Greater than operation from the stack.
     */
    void opGt() override;

    /**
     * Greater equal operation from the stack.
     */
    void opGe() override;

    /**
     * Add number to the stack.
     * @param val - number
     */
    void valNumber(double val) override;

    /**
     * Add string to the stack.
     * @param val - string
     */
    void valString(std::string val) override;

    /**
     * Add reference on cell to the stack.
     * @param val - reference on cell
     */
    void valReference(std::string val) override;

    /**
     * Add range of cells to the stack.
     * @param val - range of cells
     */
    void valRange(std::string val) override;

    /**
     * Add function call to the stack.
     * @param fnName - function name
     * @param paramCount - count of parameters
     */
    void funcCall(std::string fnName, int paramCount) override; // function call

    /**
     * Get the stack of operations.
     * @return - stack of operations
     */
    std::deque<std::shared_ptr<COperation>> getStack();

private:
    /**
     * Stack of operations.
     */
    std::deque<std::shared_ptr<COperation>> m_Stack;
};


// *—————————————————————————————————————————————————CMyExpressionBuilder.cpp——————————————————————————————————————————————————————* //

void CMyExpressionBuilder::opAdd() {
    m_Stack.push_back(std::make_shared<CAddition>());
}

void CMyExpressionBuilder::opSub() {
    m_Stack.push_back(std::make_shared<CSubtraction>());
}

void CMyExpressionBuilder::opMul() {
    m_Stack.push_back(std::make_shared<CMultiplication>());
}

void CMyExpressionBuilder::opDiv() {
    m_Stack.push_back(std::make_shared<CDivision>());
}

void CMyExpressionBuilder::opPow() {
    m_Stack.push_back(std::make_shared<CPower>());
}

void CMyExpressionBuilder::opNeg() {
    m_Stack.push_back(std::make_shared<CNegation>());
}

void CMyExpressionBuilder::opEq() {
    m_Stack.push_back(std::make_shared<CEqual>());
}

void CMyExpressionBuilder::opNe() {
    m_Stack.push_back(std::make_shared<CNotEqual>());
}

void CMyExpressionBuilder::opLt() {
    m_Stack.push_back(std::make_shared<CLessThan>());
}

void CMyExpressionBuilder::opLe() {
    m_Stack.push_back(std::make_shared<CLessEqual>());
}

void CMyExpressionBuilder::opGt() {
    m_Stack.push_back(std::make_shared<CGreaterThan>());
}

void CMyExpressionBuilder::opGe() {
    m_Stack.push_back(std::make_shared<CGreaterEqual>());
}

void CMyExpressionBuilder::valNumber(double val) {
    m_Stack.push_back(std::make_shared<CNumber>(val));
}

void CMyExpressionBuilder::valString(std::string val) {
    m_Stack.push_back(std::make_shared<CString>(val));
}

void CMyExpressionBuilder::valReference(std::string val) {
    m_Stack.push_back(std::make_shared<CReference>(val));
}

void CMyExpressionBuilder::valRange(std::string val) {
    m_Stack.push_back(std::make_shared<CValRange>());
}

void CMyExpressionBuilder::funcCall(std::string fnName, int paramCount) {
    m_Stack.push_back(std::make_shared<CFuncCall>());
}

std::deque<std::shared_ptr<COperation>> CMyExpressionBuilder::getStack() {
    return m_Stack;
}

// *—————————————————————————————————————————————————CSpreadsheet.h——————————————————————————————————————————————————————* //

/**
 * Class representing a spreadsheet.
 */
class CSpreadsheet {
public:
    static unsigned capabilities() {
        return SPREADSHEET_CYCLIC_DEPS | SPREADSHEET_FUNCTIONS | SPREADSHEET_FILE_IO | SPREADSHEET_SPEED
            /*| SPREADSHEET_PARSER*/;
    }

    CSpreadsheet();

    /**
     * Load the spreadsheet from the input stream.
     * @param is - input stream
     * @return
     */
    bool load(std::istream &is);

    /**
     * Save the spreadsheet to the output stream.
     * @param os - output stream
     * @return
     */
    bool save(std::ostream &os) const;

    /**
     * Set the contents of the cell.
     * @param pos - position of the cell
     * @param contents - contents of the cell
     * @return - true if the cell was set successfully
     */
    bool setCell(CPos pos, std::string contents);

    /**
     * Get the value of the cell.
     * @param pos - position of the cell
     * @return - value of the cell
     */
    CValue getValue(CPos pos);

    /**
     * Copy a rectangle of cells from one position to another.
     * @param dst - destination position
     * @param src - source position
     * @param w  - width
     * @param h  - height
     */
    void copyRect(CPos dst, CPos src, int w = 1, int h = 1);

private:
    /**
     * Map of cells.
     */
    std::map<CPos, CCell> m_Sheet;
};

// *—————————————————————————————————————————————————CSpreadsheet.cpp——————————————————————————————————————————————————————* //

CSpreadsheet::CSpreadsheet() {}

bool CSpreadsheet::load(std::istream &is) {
    std::map < CPos, CCell > newSheet;
    size_t size;
    if (!is.read(reinterpret_cast<char *>(&size), sizeof(size))) return false;
    for (size_t i = 0; i < size; i++) {
        CPos pos;
        CCell cell;
        if (!pos.loadBinary(is) || !cell.loadBinary(is)) return false;
        newSheet[pos] = cell;
    }
    m_Sheet = std::move(newSheet);
    return true;
}

bool CSpreadsheet::save(std::ostream &os) const {
    auto size = m_Sheet.size();
    os.write(reinterpret_cast<const char *>(&size), sizeof(size));
    for (const auto &[pos, cell]: m_Sheet) {
        if (!pos.saveBinary(os)) return false; // Serialize position
        if (!cell.saveBinary(os)) return false; // Serialize cell contents
    }
    return true;
}

bool CSpreadsheet::setCell(CPos pos, std::string contents) {
    // Check for formula (starts with '=')
    if (contents.starts_with('=')) {
        try {
            CMyExpressionBuilder builder;
            parseExpression(contents, builder);
            m_Sheet[pos].m_Stack = builder.getStack();
        } catch (const std::exception &e) {
            std::cout << "Invalid formula" << std::endl;
            return false;
        }
    } else {
        try {
            // Check if the content is numeric
            size_t idx;
            double numericValue = std::stod(contents, &idx);
            if (idx == contents.length()) {
                // Store as a number
                m_Sheet[pos].m_Stack.push_back(std::make_shared<CNumber>(numericValue));
            } else {
                // Store as a string
                m_Sheet[pos].m_Stack.push_back(std::make_shared<CString>(contents));
            }
        } catch (const std::exception &e) {
            // Store as a string
            m_Sheet[pos].m_Stack.push_back(std::make_shared<CString>(contents));
        }
    }
    return true;
}

CValue CSpreadsheet::getValue(CPos pos) {
    // Check if the cell exists in the map
    auto it = m_Sheet.find(pos);
    if (it != m_Sheet.end() && !it->second.m_Stack.empty()) {
//            m_Sheet[pos].m_IsCalculated = true;
//        std::cout << "Calculating...\n";
        return it->second.calculateCell(m_Sheet);
    }
    // Return undefined if the cell does not exist
    return std::monostate{};
}

void CSpreadsheet::copyRect(CPos dst, CPos src, int w, int h) {

    // Calculate offset between source and destination
    int rowOffset = dst.m_Row - src.m_Row;
    int columnOffset = dst.m_Column - src.m_Column;

    std::map < CPos, CCell > newSheet;
    // Copy cells from source to destination
    for (int x = 0; x < w; x++) {
        for (int y = 0; y < h; y++) {
            CPos srcPos = {src.m_Row + y, src.m_Column + x};
            CPos dstPos = {dst.m_Row + y, dst.m_Column + x};

            // Check if the source cell exists
            auto it = m_Sheet.find(srcPos);
            if (it != m_Sheet.end()) {
                // Copy the cell to the destination
                newSheet[dstPos].m_Stack.clear();
                for (const auto &operation: it->second.m_Stack) {
                    newSheet[dstPos].m_Stack.push_back(operation->clone());
                }

                // Update the cell reference in the formula
                for (auto &operation: newSheet[dstPos].m_Stack)
                    if (auto reference = std::dynamic_pointer_cast<CReference>(operation))
                        reference->setCPos(rowOffset, columnOffset);
            } else {
                // Clear the destination cell if the source cell does not exist
                newSheet[dstPos].m_Stack.clear();
            }
        }
    }

    //Iterate map and move cells to the original map
    for (auto &[pos, cell]: newSheet) {
        m_Sheet[pos] = std::move(cell);
    }

}

// *—————————————————————————————————————————————————CReference.cpp————————————————————————————————————————————————* //

CReference::CReference(std::string &str) : m_Pos(str) {}

CValue
CReference::evaluate(std::deque<std::shared_ptr<COperation>> &stack, std::map<CPos, CCell> &sheet, int &depth) const {
    // Check if the cell exists in the map
    depth++;
    auto it = sheet.find(m_Pos);
    if (it != sheet.end() && !it->second.m_Stack.empty()) {
//        std::cout << "Calculating...\n";
        return it->second.calculateCell(sheet);
    }
    // Return undefined if the cell does not exist
    return std::monostate{};
}

CPos CReference::getCPos() {
    return m_Pos;
}

void CReference::setCPos(int rowOffset, int columnOffset) {
    if (!m_Pos.m_AbsRow)
        m_Pos.m_Row += rowOffset;
    if (!m_Pos.m_AbsColumn)
        m_Pos.m_Column += columnOffset;
}

std::shared_ptr<COperation> CReference::clone() const {
    return std::make_shared<CReference>(*this);
}

bool CReference::saveBinary(std::ostream &os) const {
    m_Pos.saveBinary(os);
    return os.good();
}

bool CReference::loadBinary(std::istream &is) {
    return m_Pos.loadBinary(is);
}

int CReference::getTypeId() const {
    return 13;
}

// *—————————————————————————————————————————————————COperation.cpp————————————————————————————————————————————* //

std::shared_ptr<COperation> COperation::createOperationFromType(int typeId) {
    switch (typeId) {
        case 1:
            return std::make_shared<CAddition>();
        case 2:
            return std::make_shared<CSubtraction>();
        case 3:
            return std::make_shared<CMultiplication>();
        case 4:
            return std::make_shared<CDivision>();
        case 5:
            return std::make_shared<CPower>();
        case 6:
            return std::make_shared<CNegation>();
        case 7:
            return std::make_shared<CEqual>();
        case 8:
            return std::make_shared<CNotEqual>();
        case 9:
            return std::make_shared<CLessThan>();
        case 10:
            return std::make_shared<CLessEqual>();
        case 11:
            return std::make_shared<CGreaterThan>();
        case 12:
            return std::make_shared<CGreaterEqual>();
        case 13:
            return std::make_shared<CReference>();
        case 14:
            return std::make_shared<CNumber>();
        case 15:
            return std::make_shared<CString>();
        case 16:
            return std::make_shared<CValRange>();
        case 17:
            return std::make_shared<CFuncCall>();
        default:
            return nullptr;
    }

}

// *—————————————————————————————————————————————————CCell.cpp————————————————————————————————————————————* //
bool CCell::loadBinary(std::istream &is) {
    size_t stackSize;
    is.read(reinterpret_cast<char *>(&stackSize), sizeof(stackSize));
    m_Stack.clear();

    for (size_t i = 0; i < stackSize; ++i) {
        int typeId;
        is.read(reinterpret_cast<char *>(&typeId), sizeof(typeId));

        auto op = COperation::createOperationFromType(typeId);
        if (!op || !op->loadBinary(is)) return false;

        m_Stack.push_back(op);
    }

    is.read(reinterpret_cast<char *>(&m_IsCalculated), sizeof(m_IsCalculated));

    return true;
}

// *—————————————————————————————————————————————————PROGTEST——————————————————————————————————————————————————————* //

#ifndef __PROGTEST__

bool valueMatch(const CValue &r,
                const CValue &s) {
    if (r.index() != s.index())
        return false;
    if (r.index() == 0)
        return true;
    if (r.index() == 2)
        return std::get<std::string>(r) == std::get<std::string>(s);
    if (std::isnan(std::get<double>(r)) && std::isnan(std::get<double>(s)))
        return true;
    if (std::isinf(std::get<double>(r)) && std::isinf(std::get<double>(s)))
        return (std::get<double>(r) < 0 && std::get<double>(s) < 0)
               || (std::get<double>(r) > 0 && std::get<double>(s) > 0);
    return fabs(std::get<double>(r) - std::get<double>(s)) <= 1e8 * DBL_EPSILON * fabs(std::get<double>(r));
}

int main() {
// *—————————————————————————————————————————————————My Tests——————————————————————————————————————————————————————* //


    // SetCell and GetValue for numbers and strings
    CSpreadsheet ss;

    ss.setCell(CPos("G321"), "");
    assert(valueMatch(ss.getValue(CPos("G321")), ""));

    ss.setCell(CPos("A1"), "10");
    CValue val = ss.getValue(CPos("A1"));
    assert(valueMatch(val, CValue(10.0)));

    ss.setCell(CPos("C2"), "20.5");
    val = ss.getValue(CPos("C2"));
    assert(valueMatch(val, CValue(20.5)));

    ss.setCell(CPos("B33"), "3e1");
    val = ss.getValue(CPos("B33"));
    assert(valueMatch(val, CValue(30.0)));

    ss.setCell(CPos("G4"), "Hello World");
    val = ss.getValue(CPos("G4"));
    assert(valueMatch(val, CValue("Hello World")));

    ss.setCell(CPos("A1"), "2");
    ss.setCell(CPos("B2"), "4");
    ss.setCell(CPos("I4"), "=A1 + B2 * 3");
    assert(valueMatch(ss.getValue(CPos("I4")), CValue(14.0)));

// *—————————————————————————————————————————————————Progtest Tests——————————————————————————————————————————————————————* //

    CSpreadsheet x0, x1;
    std::ostringstream oss;
    std::istringstream iss;
    std::string data;

    assert (x0.setCell(CPos("A1"), "10"));
    assert (x0.setCell(CPos("A2"), "20.5"));
    assert (x0.setCell(CPos("A3"), "3e1"));
    assert (x0.setCell(CPos("A4"), "=40"));
    assert (x0.setCell(CPos("A5"), "=5e+1"));
    assert (x0.setCell(CPos("A6"), "raw text with any characters, including a quote \" or a newline\n"));
    assert (x0.setCell(CPos("A7"),
                       "=\"quoted string, quotes must be doubled: \"\". Moreover, backslashes are needed for C++.\""));
    assert (valueMatch(x0.getValue(CPos("A1")), CValue(10.0)));
    assert (valueMatch(x0.getValue(CPos("A2")), CValue(20.5)));
    assert (valueMatch(x0.getValue(CPos("A3")), CValue(30.0)));
    assert (valueMatch(x0.getValue(CPos("A4")), CValue(40.0)));
    assert (valueMatch(x0.getValue(CPos("A5")), CValue(50.0)));
    assert (valueMatch(x0.getValue(CPos("A6")),
                       CValue("raw text with any characters, including a quote \" or a newline\n")));
    assert (valueMatch(x0.getValue(CPos("A7")),
                       CValue("quoted string, quotes must be doubled: \". Moreover, backslashes are needed for C++.")));
    assert (valueMatch(x0.getValue(CPos("A8")), CValue()));
    assert (valueMatch(x0.getValue(CPos("AAAA9999")), CValue()));
    assert (x0.setCell(CPos("B1"), "=A1+A2*A3"));
    assert (x0.setCell(CPos("B2"), "= -A1 ^ 2 - A2 / 2   "));
    assert (x0.setCell(CPos("B3"), "= 2 ^ $A$1"));
    assert (x0.setCell(CPos("B4"), "=($A1+A$2)^2"));
    assert (x0.setCell(CPos("B5"), "=B1+B2+B3+B4"));
    assert (x0.setCell(CPos("B6"), "=B1+B2+B3+B4+B5"));
    assert (valueMatch(x0.getValue(CPos("B1")), CValue(625.0)));
    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-110.25)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(1024.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(930.25)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(2469.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(4938.0)));
    assert (x0.setCell(CPos("A1"), "12"));
    assert (valueMatch(x0.getValue(CPos("B1")), CValue(627.0)));
    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-154.25)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(1056.25)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(5625.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(11250.0)));
    x1 = x0;
    assert (x0.setCell(CPos("A2"), "100"));
    assert (x1.setCell(CPos("A2"), "=A3+A5+A4"));
    assert (valueMatch(x0.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(38916.0)));
    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3612.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-204.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(17424.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(24928.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(49856.0)));
    oss.clear();
    oss.str("");
    assert (x0.save(oss));
    data = oss.str();
    iss.clear();
    iss.str(data);
    assert (x1.load(iss));
    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
    assert (x0.setCell(CPos("A3"), "4e1"));
    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
    oss.clear();
    oss.str("");
    assert (x0.save(oss));
    data = oss.str();
    for (size_t i = 0; i < std::min<size_t>(data.length(), 10); i++)
        data[i] ^= 0x5a;
    iss.clear();
    iss.str(data);
    assert (!x1.load(iss));
    assert (x0.setCell(CPos("D0"), "10"));
    assert (x0.setCell(CPos("D1"), "20"));
    assert (x0.setCell(CPos("D2"), "30"));
    assert (x0.setCell(CPos("D3"), "40"));
    assert (x0.setCell(CPos("D4"), "50"));
    assert (x0.setCell(CPos("E0"), "60"));
    assert (x0.setCell(CPos("E1"), "70"));
    assert (x0.setCell(CPos("E2"), "80"));
    assert (x0.setCell(CPos("E3"), "90"));
    assert (x0.setCell(CPos("E4"), "100"));
    assert (x0.setCell(CPos("F10"), "=D0+5"));
    assert (x0.setCell(CPos("F11"), "=$D0+5"));
    assert (x0.setCell(CPos("F12"), "=D$0+5"));
    assert (x0.setCell(CPos("F13"), "=$D$0+5"));
    x0.copyRect(CPos("G11"), CPos("F10"), 1, 4);
    assert (valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F14")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G10")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
    assert (valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
    assert (valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));
    x0.copyRect(CPos("G11"), CPos("F10"), 2, 4);
    assert (valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F14")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G10")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
    assert (valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
    assert (valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("H10")), CValue()));
    assert (valueMatch(x0.getValue(CPos("H11")), CValue()));
    assert (valueMatch(x0.getValue(CPos("H12")), CValue()));
    assert (valueMatch(x0.getValue(CPos("H13")), CValue(35.0)));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue()));
    assert (x0.setCell(CPos("F0"), "-27"));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));
    x0.copyRect(CPos("H12"), CPos("H13"), 1, 2);
    assert (valueMatch(x0.getValue(CPos("H12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("H13")), CValue(-22.0)));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));
    return EXIT_SUCCESS;
}

#endif /* __PROGTEST__ */