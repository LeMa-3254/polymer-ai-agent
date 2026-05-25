#!/usr/bin/env python3
"""Weekly polymer AI digest generator — 2026-05-25"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-05-25"

data = {
    "date": DATE,
    "headline": (
        "Foundation models, million-molecule datasets, and real-time process AI are "
        "converging to make 2026 the breakout year for AI-driven polymer science."
    ),
    "executive_summary": (
        "The week of 25 May 2026 crystallises a year of rapid progress across every "
        "dimension of AI-driven polymer science. Georgia Tech's POLYT5—the first "
        "encoder-decoder foundation chemical language model for generative polymer design—"
        "has now been experimentally validated on a novel dielectric material, proving "
        "that AI-generated polymer candidates can survive the crucible of wet-lab "
        "synthesis. Simultaneously, LLNL and Meta released the Open Polymers 2026 "
        "(OPoly26) dataset containing 6.35 million DFT-quality calculations, the largest "
        "and most chemically diverse quantum-chemistry dataset ever assembled for "
        "polymeric systems.\n\n"
        "On the property-prediction front, multimodal architectures are setting new "
        "benchmarks. PolyLLMem fuses Llama 3 text embeddings with Uni-Mol 3D structural "
        "embeddings via LoRA fine-tuning, achieving state-of-the-art accuracy on glass "
        "transition temperature, dielectric constant, and bandgap. Separately, an adaptive "
        "gating mechanism for GNN-fingerprint fusion in Macromolecules 2026 has been "
        "incorporated into the Polymer Genome prediction pipeline, further extending its "
        "multi-task coverage. LLMs are also proving their worth as data-extraction "
        "engines: automated pipelines now mine over one million polymer property records "
        "from 681,000 journal articles, rapidly closing the data-scarcity gap.\n\n"
        "Characterisation automation is accelerating. Deep learning models now predict "
        "FTIR fingerprint-region spectra for six major commodity polymers with up to 87% "
        "accuracy, while deep residual networks classify pristine polymer types at 97.9% "
        "accuracy from raw spectral data. The PlasticAnalytics platform packages these "
        "capabilities into a deployable suite for environmental microplastics monitoring. "
        "AI-aided crystallisation elution fractionation (CEF) is maturing as a rapid, "
        "automated composition tool for polyolefin quality control.\n\n"
        "Industrial deployment is delivering documented returns. AI process-control "
        "platforms such as Imubit report more than 2% reduction in off-spec production "
        "and 10–20% cuts in natural-gas consumption across polymer manufacturing lines. "
        "For sustainability, multi-objective AI frameworks for plastic pyrolysis are "
        "navigating complex Pareto trade-offs to improve oil-product recovery by 15%, "
        "and ensemble deep-learning classifiers are enabling 98%-accurate automated "
        "polymer sorting for recycling streams—a capability that is already attracting "
        "industrial investment ahead of tightening EU plastic-waste regulations."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: Foundation Language Model Generates Experimentally Validated Dielectric Polymers",
            "description": (
                "Georgia Tech's Ramprasad Group published POLYT5 in npj Artificial Intelligence "
                "(March 2026), an encoder-decoder T5-based foundation model for generative "
                "polymer design. The model accepts desired property targets as natural-language "
                "prompts and outputs chemically valid polymer SMILES, generating 18,000+ "
                "dielectric candidates of which one was physically synthesised and confirmed to "
                "meet predictions. This is the first end-to-end experimental validation of a "
                "generative polymer foundation model, establishing a new benchmark for the field."
            ),
            "source": "https://www.nature.com/articles/s44387-026-00087-1",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26: LLNL and Meta Release 6.35 Million DFT Polymer Calculations",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset—a joint effort by Lawrence Livermore "
                "National Laboratory and Meta—contains 6.35 million DFT-accuracy quantum "
                "chemistry calculations spanning monomer composition, chain architecture, degree "
                "of polymerization, and solvation environments. It is the largest and most "
                "chemically diverse polymer quantum-chemistry corpus ever made publicly available. "
                "Models trained on OPoly26 predict electronic, thermal, and mechanical properties "
                "from SMILES in milliseconds, dramatically reducing laboratory synthesis needs "
                "during early-stage screening."
            ),
            "source": "https://arxiv.org/pdf/2512.23117",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem Fuses Llama 3 and Uni-Mol for State-of-the-Art Property Prediction",
            "description": (
                "Published in Chemistry of Materials 2026, PolyLLMem integrates text embeddings "
                "from Llama 3 with 3D molecular structure embeddings from Uni-Mol via Low-Rank "
                "Adaptation (LoRA) fine-tuning. The multimodal fusion achieves state-of-the-art "
                "accuracy on glass transition temperature, dielectric constant, and bandgap "
                "benchmarks. The work confirms that general-purpose LLMs encode transferable "
                "chemical knowledge that substantially boosts performance on polymer-specific "
                "tasks without expensive domain-specific pretraining from scratch."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "LLMs Extract 1 Million Polymer Property Records from 681K Scientific Articles",
            "description": (
                "A Nature Communications Materials study demonstrated automated extraction of "
                "over one million polymer property records—covering 24 distinct properties for "
                "more than 106,000 unique polymers—from a 2.4-million-article full-text corpus "
                "using large language models. The pipeline handles chemical entity recognition, "
                "unit normalization, and confidence scoring autonomously. This capability "
                "effectively mines decades of legacy literature into structured ML-ready data, "
                "directly feeding polymer informatics platforms and training pipelines at a "
                "scale impossible with manual curation."
            ),
            "source": "https://www.nature.com/articles/s43246-024-00708-9",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Adaptive GNN-Fingerprint Fusion Sets New Benchmark for Thermal Properties",
            "description": (
                "A 2026 Macromolecules paper introduced an adaptive gating mechanism that "
                "dynamically fuses graph neural network embeddings with classical molecular "
                "fingerprints for multi-task prediction of polymer thermal properties—Tg, Tm, "
                "and thermal conductivity. The gating layer weights the two representation "
                "streams per property and per polymer class, outperforming single-representation "
                "baselines on five benchmark datasets. The approach has since been incorporated "
                "into the Polymer Genome online prediction platform."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.macromol.3c02401",
            "impact": "medium",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning Predicts Polymer FTIR Fingerprint Region at 87% Accuracy",
            "description": (
                "A 2026 MRS Advances study evaluated multilayer perceptron and deep residual "
                "network architectures for predicting FTIR fingerprint-region spectra of PET, "
                "HDPE, PVC, LDPE, PP, and PS. The MLP achieved 87% accuracy for polystyrene "
                "and 84% for HDPE, while a deep residual network reached 97.9% classification "
                "accuracy on pristine polymer spectra. A parallel MELP ensemble model for "
                "micro- and nanoplastic identification demonstrated superior generalizability "
                "compared to existing deep-neural-network baselines."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Processing",
            "title": "Imubit AI Platform Cuts Polymer Plant Energy 10–20% with Real-Time ML Control",
            "description": (
                "Industrial AI platform Imubit published documented results from polymer "
                "manufacturing deployments: real-time ML process control reduced off-spec "
                "production by over 2%, cut natural-gas consumption by 10–20%, and raised "
                "throughput by 1–3%—all without capital equipment changes. The system "
                "dynamically adjusts heat, pressure, and flow using reinforcement-learning-style "
                "optimization with millisecond response times. These are among the first large-"
                "scale, independently verified AI efficiency gains published for continuous "
                "polymer manufacturing operations."
            ),
            "source": "https://imubit.com/articles/polymer-processing",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "Multi-Objective AI Optimization Improves Plastic Pyrolysis Product Recovery by 15%",
            "description": (
                "A 2026 Polymers study deployed a surrogate-assisted multi-objective "
                "optimization framework integrating an energy return-on-investment metric to "
                "jointly maximize oil-product yield, minimize energy inputs, and reduce CO2 "
                "emissions in plastic waste pyrolysis. The AI-navigated Pareto-optimal process "
                "conditions improved product recovery by 15% over conventional single-objective "
                "baselines. The framework is designed for direct integration with industrial "
                "pyrolysis reactor control systems."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "medium",
        },
        {
            "category": "Generative Design",
            "title": "Conditional GAN Delivers Wet-Lab Validated Biomedical Polymer Coatings",
            "description": (
                "Published in RSC Digital Discovery 2026, a constraint-aware conditional GAN "
                "framework generates polymer coating compositions conditioned on desired "
                "biological performance targets—anti-fouling and cell-adhesion properties for "
                "medical implant surfaces. The AI-generated formulations were experimentally "
                "validated, meeting biological targets that manual optimization required months "
                "to achieve. This represents one of the first wet-lab-confirmed inverse designs "
                "for a functional polymer coating system."
            ),
            "source": "https://pubs.rsc.org/en/content/articlehtml/2026/dd/d5dd00332f",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "Ensemble Deep Learning Achieves 98% Accuracy for Automated Polymer Sorting",
            "description": (
                "A Polymers 2026 study combined ensemble deep learning with heuristic "
                "optimization (genetic algorithms and particle swarm) to classify polymer "
                "types from spectral and imaging data at 98% accuracy—a 40% reduction in "
                "error rate versus prior CNN-only approaches on challenging mixed-polymer "
                "pellets. The system targets integration into conveyor-based recycling "
                "sorting lines, enabling real-time stream separation at throughputs "
                "impractical for human operators, and is already drawing industrial "
                "investment ahead of tightening EU plastic-waste regulations."
            ),
            "source": "https://doi.org/10.3390/polym18101208",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "Polymer foundation models (POLYT5, PolyLLMem) are replacing generic ML for materials tasks, enabling natural-language design",
        "Massive open datasets (OPoly26, POINT², OpenPoly) are closing the data-scarcity bottleneck that limited earlier polymer ML",
        "Inverse design is maturing from in silico to wet-lab validated: generative models now produce experimentally confirmed structures",
        "Industrial process AI is delivering documented ROI — 10–20% energy savings and 2%+ quality gains in polymer manufacturing",
        "LLM-powered literature mining is converting decades of legacy publications into structured, ML-ready polymer property databases",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Georgia Tech encoder-decoder foundation model for generative polymer design; "
                "accepts property targets as text prompts, outputs valid SMILES (npj AI, 2026)."
            ),
            "url": "https://www.nature.com/articles/s44387-026-00087-1",
        },
        {
            "name": "OPoly26 Dataset",
            "purpose": (
                "LLNL + Meta open dataset of 6.35M DFT-accuracy polymer calculations; "
                "backbone training corpus for AI property and generative models (arXiv, 2025)."
            ),
            "url": "https://arxiv.org/abs/2512.23117",
        },
        {
            "name": "Polymer Genome",
            "purpose": (
                "Web platform for ML-based polymer property prediction covering Tg, bandgap, "
                "dielectric constant and more; now incorporates GNN-fingerprint fusion models."
            ),
            "url": "https://www.polymergenome.org",
        },
        {
            "name": "PlasticAnalytics",
            "purpose": (
                "Deep-learning-powered spectral library and analytical suite for FTIR-based "
                "polymer and microplastic identification (ACS Environmental Science & Technology)."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.est.6c01309",
        },
        {
            "name": "PolyLLMem",
            "purpose": (
                "Multimodal model fusing Llama 3 text embeddings with Uni-Mol 3D structure "
                "embeddings for polymer property prediction via LoRA fine-tuning (Chem. Mater. 2026)."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
        },
        {
            "name": "Imubit Process AI",
            "purpose": (
                "Industrial ML platform for real-time polymer manufacturing optimization; "
                "documented 10–20% energy savings and 2%+ reduction in off-spec production."
            ),
            "url": "https://imubit.com/articles/polymer-processing",
        },
    ],
    "outlook": (
        "The next 12 months will see polymer foundation models become standard infrastructure, "
        "mirroring the trajectory of protein language models such as ESM and AlphaFold in "
        "biology. POLYT5 and its successors will be fine-tuned for specialised sub-domains—"
        "high-temperature dielectrics, bioresorbable scaffolds, conductive elastomers—enabling "
        "on-demand inverse design that was science fiction five years ago. The integration of "
        "AI with high-throughput robotic synthesis platforms will close the design-synthesize-"
        "test loop autonomously, compressing materials development timelines from years to "
        "weeks. On the sustainability front, AI-enabled sorting and depolymerisation tools will "
        "become central to chemical recycling economics, directly influencing corporate "
        "sustainability commitments and EU regulatory compliance strategies. The main challenge "
        "ahead is not capability but deployment: bridging the gap between academic models and "
        "production-grade tools that polymer engineers can trust and integrate into existing "
        "workflows without requiring deep ML expertise."
    ),
}

# ── Generate documents ─────────────────────────────────────────────────────

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print("Generating Word document...")
    generate_word(data, docx_path)

    print("Generating PowerPoint...")
    generate_pptx(data, pptx_path)

    print(f"\nDone! Files written to {output_dir.resolve()}")
    print(f"  {docx_path.name}")
    print(f"  {pptx_path.name}")
