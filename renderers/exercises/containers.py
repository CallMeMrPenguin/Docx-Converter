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

        old_sign_w = getattr(self, "current_group_sign_pic_w_cm", None)
        old_sign_h = getattr(self, "current_group_sign_pic_h_cm", None)
        old_sign_gap = getattr(self, "current_group_sign_pic_gap_cm", None)

        if hasattr(self, "compute_group_sign_mcq_params"):
            s_w, s_h, s_gap = self.compute_group_sign_mcq_params(block.children, printable_width_cm)
            if s_w is not None:
                self.current_group_sign_pic_w_cm = s_w
                self.current_group_sign_pic_h_cm = s_h
                self.current_group_sign_pic_gap_cm = s_gap

        old_inside_num = getattr(self, "is_inside_num_container", False)
        self.is_inside_num_container = True
        try:
            if hasattr(self, "render"):
                self.render(block.children, doc, word, is_root=False)
        finally:
            self.is_inside_num_container = old_inside_num
            self.current_group_opt_cols, self.current_group_max_item_len = old_cols, old_len
            self.current_group_sign_pic_w_cm = old_sign_w
            self.current_group_sign_pic_h_cm = old_sign_h
            self.current_group_sign_pic_gap_cm = old_sign_gap
