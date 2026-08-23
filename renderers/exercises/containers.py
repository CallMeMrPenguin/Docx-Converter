from uln_parser import ULNBlock


class ContainersRendererMixin:
    """Renders hierarchical structural containers such as [NUM] ... [/NUM]."""

    def render_num_container(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders auto-numbered container [NUM] ... [/NUM].
        Flags the first question in this section to start a new independent list.
        Pre-computes uniform option alignment across all questions in the exercise.
        """
        if not block.children:
            return

        self.is_first_question_in_num_block = True

        opt_blocks = [c for c in block.children if c.tag == "OPT"]
        old_cols = getattr(self, "current_group_opt_cols", None)
        old_len = getattr(self, "current_group_max_item_len", None)

        if opt_blocks and hasattr(self, "compute_group_option_params"):
            g_cols, g_len = self.compute_group_option_params(opt_blocks, printable_width_cm)
            self.current_group_opt_cols = g_cols
            self.current_group_max_item_len = g_len

        try:
            if hasattr(self, "render"):
                self.render(block.children, doc, word, is_root=False)
        finally:
            self.current_group_opt_cols, self.current_group_max_item_len = old_cols, old_len
