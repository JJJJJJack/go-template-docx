package template

import (
	"text/template"
	"text/template/parse"
)

// collectVariables walks the parse tree and collects variables
func collectVariables(node parse.Node, vars map[string]struct{}) {
	switch n := node.(type) {
	case *parse.ListNode:
		if n != nil {
			for _, elem := range n.Nodes {
				collectVariables(elem, vars)
			}
		}
	case *parse.ActionNode:
		collectVariables(n.Pipe, vars)
	case *parse.RangeNode:
		collectVariables(n.Pipe, vars)
		collectVariables(n.List, vars)
		collectVariables(n.ElseList, vars)
	case *parse.IfNode:
		collectVariables(n.Pipe, vars)
		collectVariables(n.List, vars)
		collectVariables(n.ElseList, vars)
	case *parse.WithNode:
		collectVariables(n.Pipe, vars)
		collectVariables(n.List, vars)
		collectVariables(n.ElseList, vars)
	case *parse.TemplateNode:
		// reference to another template, but variables come from that template's Tree
	case *parse.PipeNode:
		for _, cmd := range n.Cmds {
			collectVariables(cmd, vars)
		}
	case *parse.CommandNode:
		for _, arg := range n.Args {
			collectVariables(arg, vars)
		}
	case *parse.FieldNode:
		// Example: {{.User.Name}}
		vars[n.String()] = struct{}{}
	case *parse.VariableNode:
		// Example: {{$var}} or {{$var.Field}}
		vars[n.String()] = struct{}{}
	}
}

// ExtractAllVariables extract variables from ALL templates in a set
func ExtractAllVariables(t *template.Template) map[string]struct{} {
	vars := make(map[string]struct{})
	for _, tpl := range t.Templates() {
		if tpl.Tree != nil && tpl.Tree.Root != nil {
			collectVariables(tpl.Tree.Root, vars)
		}
	}

	return vars
}
