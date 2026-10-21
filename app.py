                    st.dataframe(pd.DataFrame(st.session_state["sains_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Sainsbury's workbook",
                    data=st.session_state["sains_report_bytes"],
                    file_name="sainsburys_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_sainsbury",
                )

        if waitrose_template:
            if st.button("Build Waitrose report"):
                waitrose_output, waitrose_results, waitrose_summary = build_waitrose_workbook(waitrose_template, df)
                st.session_state["waitrose_report_bytes"] = waitrose_output.getvalue()
                st.session_state["waitrose_report_results"] = waitrose_results
                st.session_state["waitrose_summary"] = waitrose_summary
            if "waitrose_report_results" in st.session_state:
                st.subheader("Waitrose processing summary")
                st.dataframe(pd.DataFrame(st.session_state["waitrose_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("waitrose_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["waitrose_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Waitrose workbook",
                    data=st.session_state["waitrose_report_bytes"],
                    file_name="waitrose_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_waitrose",
                )

        if coop_template:
            if st.button("Build The Co-operative report"):
                coop_output, coop_results, coop_summary = build_coop_workbook(coop_template, df)
                st.session_state["coop_report_bytes"] = coop_output.getvalue()
                st.session_state["coop_report_results"] = coop_results
                st.session_state["coop_summary"] = coop_summary
            if "coop_report_results" in st.session_state:
                st.subheader("The Co-operative processing summary")
                st.dataframe(pd.DataFrame(st.session_state["coop_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("coop_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["coop_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated The Co-operative workbook",
                    data=st.session_state["coop_report_bytes"],
                    file_name="cooperative_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_coop",
                )

        if tesco_template:
            if st.button("Build Tesco report"):
                tesco_output, tesco_results, tesco_summary = build_tesco_workbook(tesco_template, df)
                st.session_state["tesco_report_bytes"] = tesco_output.getvalue()
                st.session_state["tesco_report_results"] = tesco_results
                st.session_state["tesco_summary"] = tesco_summary
            if "tesco_report_results" in st.session_state:
                st.subheader("Tesco processing summary")
                st.dataframe(pd.DataFrame(st.session_state["tesco_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("tesco_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["tesco_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Tesco workbook",
                    data=st.session_state["tesco_report_bytes"],
                    file_name="tesco_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_tesco",
                )

        st.divider()
        st.subheader("Post-build audit checks")
        st.caption("Highlighted rows are source audits not assigned by the active report rules. Compliance Audit records are intentionally excluded. E-Cig - No ID records still need the Indy fallback applied after checking all other client rules; they must not be treated as allocated until they appear in the generated Indy workbook.")
        unallocated, overlapping = allocation_audit(
            df, template_file, sainsbury_template, tesco_template, coop_template, waitrose_template
        )
        compliance_excluded = df.loc[df["item_to_order"].astype(str).str.strip().str.casefold().eq("compliance audit")].copy()
        m1, m2, m3 = st.columns(3)
        m1.metric("Audits not added", len(unallocated))
        m2.metric("Audits matching multiple reports", len(overlapping))
        m3.metric("Compliance Audit intentionally excluded", len(compliance_excluded))
        if not unallocated.empty:
            st.warning("Review these audits: they did not match any of the uploaded report rules and will not be added to a workbook.")
            preview_cols = [c for c in ["internal_id", "date_of_visit", "item_to_order", "site_name", "site_code", "primary_result"] if c in unallocated.columns]
            highlighted = unallocated[preview_cols].style.apply(lambda row: ["background-color: #ffe0e0; color: #7f1d1d"] * len(row), axis=1)
            st.dataframe(highlighted, use_container_width=True, hide_index=True)
            st.download_button(
                "Download unallocated audits CSV",
                data=unallocated.to_csv(index=False).encode("utf-8-sig"),
                file_name="unallocated_audits.csv",
                mime="text/csv",
                key="download_unallocated",
            )
        if not compliance_excluded.empty:
            with st.expander(f"Intentionally excluded Compliance Audit records ({len(compliance_excluded)})"):
                cols = [c for c in ["internal_id", "date_of_visit", "item_to_order", "site_name", "site_code", "primary_result"] if c in compliance_excluded.columns]
                st.dataframe(compliance_excluded[cols], use_container_width=True, hide_index=True)
        if not overlapping.empty:
            st.caption("Any remaining overlaps are shown for review. Retail rules exclude online audit types and Compliance Audit.")
            overlap_cols = [c for c in ["internal_id", "date_of_visit", "item_to_order", "site_name", "site_code", "matched_reports"] if c in overlapping.columns]
            st.dataframe(overlapping[overlap_cols], use_container_width=True)
            st.download_button(
                "Download multi-match audits CSV",
                data=overlapping.to_csv(index=False).encode("utf-8-sig"),
                file_name="multi_match_audits.csv",
                mime="text/csv",
                key="download_overlapping",
            )

    except Exception as exc:
        st.error(f"Could not process these files: {exc}")
else:
    st.info("Upload the audit export and at least one existing client workbook to enable report processing.")
