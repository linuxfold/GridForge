import Foundation

// MARK: - Worksheet conforms to CellValueProvider

extension Worksheet: CellValueProvider {}

// MARK: - Dependency Graph

/// Tracks which cells depend on which other cells, enabling efficient incremental
/// recalculation and circular reference detection.
public final class DependencyGraph {

    /// Forward map: cell -> set of cells it depends on (its precedents).
    private var precedents: [CellAddress: Set<CellAddress>] = [:]

    /// Reverse map: cell -> set of cells that depend on it (its dependents).
    private var dependentsMap: [CellAddress: Set<CellAddress>] = [:]

    public init() {}

    /// Register that `cell` depends on the given set of addresses.
    /// Replaces any previously recorded dependencies for that cell.
    public func addDependencies(cell: CellAddress, dependsOn: Set<CellAddress>) {
        // Remove old dependencies first
        removeDependencies(for: cell)

        // Record new precedents
        precedents[cell] = dependsOn

        // Update reverse map
        for dep in dependsOn {
            dependentsMap[dep, default: []].insert(cell)
        }
    }

    /// Remove all dependency information for a cell.
    public func removeDependencies(for cell: CellAddress) {
        guard let oldDeps = precedents.removeValue(forKey: cell) else { return }
        for dep in oldDeps {
            dependentsMap[dep]?.remove(cell)
            if dependentsMap[dep]?.isEmpty == true {
                dependentsMap.removeValue(forKey: dep)
            }
        }
    }

    /// Returns the set of cells that directly or transitively depend on the given cell.
    /// These are the cells that need recalculation when the given cell changes.
    public func dependents(of cell: CellAddress) -> Set<CellAddress> {
        var result = Set<CellAddress>()
        var queue = Array(directDependents(of: cell))
        while !queue.isEmpty {
            let current = queue.removeFirst()
            if result.insert(current).inserted {
                queue.append(contentsOf: directDependents(of: current))
            }
        }
        return result
    }

    /// Returns only the direct dependents (one level) of a cell.
    public func directDependents(of cell: CellAddress) -> Set<CellAddress> {
        dependentsMap[cell] ?? []
    }

    /// Returns the precedents (cells this cell depends on).
    public func precedentsOf(_ cell: CellAddress) -> Set<CellAddress> {
        precedents[cell] ?? []
    }

    /// Topological sort of cells starting from a given set, returning an evaluation
    /// order where each cell appears after all its precedents.
    /// Returns nil if a cycle is detected among the given cells.
    public func topologicalSort(from cells: Set<CellAddress>) -> [CellAddress]? {
        // Kahn's algorithm on the subgraph induced by `cells`.
        // We only consider edges within the given set.

        // Compute in-degree for each cell within the set
        var inDegree: [CellAddress: Int] = [:]
        for cell in cells {
            inDegree[cell] = 0
        }
        for cell in cells {
            let deps = precedents[cell] ?? []
            for dep in deps where cells.contains(dep) {
                inDegree[cell, default: 0] += 1
            }
        }

        // Start with cells that have no in-edges within the set
        var queue: [CellAddress] = []
        for (cell, degree) in inDegree where degree == 0 {
            queue.append(cell)
        }
        queue.sort() // deterministic order

        var sorted: [CellAddress] = []
        var index = 0
        while index < queue.count {
            let cell = queue[index]
            index += 1
            sorted.append(cell)

            // For each cell that depends on this one (within the set), decrease in-degree
            for dependent in directDependents(of: cell) where cells.contains(dependent) {
                inDegree[dependent, default: 0] -= 1
                if inDegree[dependent] == 0 {
                    queue.append(dependent)
                }
            }
        }

        // If we didn't visit all cells, there's a cycle
        guard sorted.count == cells.count else { return nil }
        return sorted
    }

    /// Detect whether adding the current dependencies for `from` would create a cycle.
    /// Returns true if a cycle is detected.
    public func detectCycle(from cell: CellAddress) -> Bool {
        // DFS from the cell's precedents to see if any path leads back to `cell`.
        let deps = precedents[cell] ?? []
        var visited = Set<CellAddress>()
        var stack = Array(deps)

        while !stack.isEmpty {
            let current = stack.removeLast()
            if current == cell { return true }
            if visited.insert(current).inserted {
                let next = precedents[current] ?? []
                stack.append(contentsOf: next)
            }
        }
        return false
    }

    /// Remove all dependency information.
    public func clear() {
        precedents.removeAll()
        dependentsMap.removeAll()
    }
}

// MARK: - Formula Engine

/// The main facade that ties tokenizing, parsing, evaluating, and dependency
/// tracking together for a spreadsheet worksheet.
public final class FormulaEngine {

    public let dependencyGraph: DependencyGraph
    public let evaluator: FormulaEvaluator

    public init() {
        self.dependencyGraph = DependencyGraph()
        self.evaluator = FormulaEvaluator()
    }

    // MARK: - Evaluate a Formula String

    /// Tokenize, parse, and evaluate a formula string (without leading "="),
    /// returning the computed CellValue.
    public func evaluate(formula: String, in worksheet: Worksheet, at address: CellAddress) -> CellValue {
        // Tokenize
        let tokenizer = Tokenizer(formula: formula)
        let tokens: [Token]
        do {
            tokens = try tokenizer.tokenize()
        } catch {
            return .error(.syntax)
        }

        // Parse
        let parser = FormulaParser(tokens: tokens)
        let ast: ASTNode
        do {
            ast = try parser.parse()
        } catch {
            return .error(.syntax)
        }

        // Update dependency graph
        let refs = extractReferences(from: ast)
        dependencyGraph.addDependencies(cell: address, dependsOn: refs)

        // Check for circular references
        if dependencyGraph.detectCycle(from: address) {
            return .error(.circular)
        }

        // Evaluate
        return evaluator.evaluate(node: ast, provider: worksheet)
    }

    // MARK: - Evaluate a Single Cell

    /// Evaluate a single cell's formula and set its value on the worksheet.
    /// Non-formula cells are left unchanged.
    public func evaluateCell(at address: CellAddress, in worksheet: Worksheet) {
        guard let cell = worksheet.cell(at: address),
              let formula = cell.formulaExpression else {
            // Not a formula cell — remove from dependency graph
            dependencyGraph.removeDependencies(for: address)
            return
        }

        let result = evaluate(formula: formula, in: worksheet, at: address)
        cell.value = result
    }

    // MARK: - Full Recalculation

    /// Evaluate ALL formula cells in the worksheet in dependency order.
    public func recalculate(worksheet: Worksheet) {
        // First pass: collect all formula cells and rebuild the dependency graph.
        dependencyGraph.clear()
        var formulaCells = Set<CellAddress>()

        for (address, cell) in worksheet.cells {
            guard let formula = cell.formulaExpression else { continue }
            formulaCells.insert(address)

            // Parse to extract references
            if let ast = parseFormula(formula) {
                let refs = extractReferences(from: ast)
                dependencyGraph.addDependencies(cell: address, dependsOn: refs)
            }
        }

        // Topological sort — if there are cycles, mark those cells as circular errors.
        guard let order = dependencyGraph.topologicalSort(from: formulaCells) else {
            // Cycle detected — evaluate what we can and mark the rest.
            evaluateWithCycleDetection(formulaCells: formulaCells, worksheet: worksheet)
            return
        }

        // Evaluate in dependency order.
        for address in order {
            guard let cell = worksheet.cell(at: address),
                  let formula = cell.formulaExpression else { continue }
            let result = evaluate(formula: formula, in: worksheet, at: address)
            cell.value = result
        }
    }

    // MARK: - Incremental Recalculation

    /// Recalculate only cells affected by a change to the given cell.
    public func recalculateAffected(changedCell: CellAddress, in worksheet: Worksheet) {
        // First, re-evaluate the changed cell itself if it's a formula.
        evaluateCell(at: changedCell, in: worksheet)

        // Find all transitive dependents.
        let affected = dependencyGraph.dependents(of: changedCell)
        guard !affected.isEmpty else { return }

        // Topological sort the affected cells.
        // Include the changed cell's dependents plus all cells they depend on that are also formula cells.
        let relevantCells = affected
        // We need to sort just the affected set in correct order.
        if let order = dependencyGraph.topologicalSort(from: relevantCells) {
            for address in order {
                guard let cell = worksheet.cell(at: address),
                      let formula = cell.formulaExpression else { continue }
                let result = evaluate(formula: formula, in: worksheet, at: address)
                cell.value = result
            }
        } else {
            // Cycle in the affected subgraph — mark cyclic cells.
            for address in relevantCells {
                guard let cell = worksheet.cell(at: address),
                      cell.isFormula else { continue }
                if dependencyGraph.detectCycle(from: address) {
                    cell.value = .error(.circular)
                } else {
                    evaluateCell(at: address, in: worksheet)
                }
            }
        }
    }

    // MARK: - Reference Extraction

    /// Extract all cell addresses referenced in an AST.
    public func extractReferences(from ast: ASTNode) -> Set<CellAddress> {
        Set(extractCellReferences(from: ast).map(\.address))
    }

    /// Extract formula references without discarding absolute flags or sheet identity.
    public func extractCellReferences(from ast: ASTNode) -> Set<CellReference> {
        var refs = Set<CellReference>()
        collectReferences(from: ast, into: &refs)
        return refs
    }

    private func collectReferences(from node: ASTNode, into refs: inout Set<CellReference>) {
        switch node {
        case .cellReference(let reference):
            refs.insert(reference)

        case .range(let start, let end):
            let range = CellRangeReference(start: start, end: end)
            refs.formUnion(range.allReferences)

        case .binaryOp(_, let left, let right):
            collectReferences(from: left, into: &refs)
            collectReferences(from: right, into: &refs)

        case .unaryOp(_, let operand):
            collectReferences(from: operand, into: &refs)

        case .functionCall(_, let args):
            for arg in args {
                collectReferences(from: arg, into: &refs)
            }

        case .number, .string, .boolean, .error:
            break
        }
    }

    // MARK: - Internal Helpers

    /// Parse a formula string into an AST, returning nil on error.
    private func parseFormula(_ formula: String) -> ASTNode? {
        let tokenizer = Tokenizer(formula: formula)
        guard let tokens = try? tokenizer.tokenize() else { return nil }
        let parser = FormulaParser(tokens: tokens)
        return try? parser.parse()
    }

    /// Evaluate formula cells with cycle detection, marking cyclic cells as errors.
    private func evaluateWithCycleDetection(formulaCells: Set<CellAddress>, worksheet: Worksheet) {
        // Identify non-cyclic cells and evaluate them first.
        var remaining = formulaCells
        var evaluated = true

        // Iteratively evaluate cells whose dependencies are all satisfied.
        while evaluated {
            evaluated = false
            for address in remaining {
                let deps = dependencyGraph.precedentsOf(address).intersection(formulaCells)
                if deps.isSubset(of: formulaCells.subtracting(remaining)) || deps.isEmpty {
                    // All precedents already evaluated or are not formula cells.
                    evaluateCell(at: address, in: worksheet)
                    remaining.remove(address)
                    evaluated = true
                }
            }
        }

        // Anything still remaining is part of a cycle.
        for address in remaining {
            worksheet.cell(at: address)?.value = .error(.circular)
        }
    }
}
